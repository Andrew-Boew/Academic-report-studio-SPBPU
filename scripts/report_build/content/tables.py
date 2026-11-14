"""Стратегия таблицы."""

from __future__ import annotations

import copy
import datetime as dt
import functools
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Cm, Pt, RGBColor
from lxml import etree

from report_build.constants import *
from report_build.errors import SpecError
from report_build.part import ReportPart
from report_build.document.cells import set_cell_margins, set_cell_no_wrap, set_cell_width, set_repeat_table_header, set_row_cant_split
from report_build.document.markup import add_bookmark, object_bookmark_name, set_run_font
from report_build.spec.textutil import clean_text, keep_previous_reference

class TableBuilder(ReportPart):
    """Таблицы отчёта."""

    def choose_widths(self, headers: list[Any], rows: list[list[Any]]) -> list[float]:
        columns = len(headers)
        weights = []
        for index in range(columns):
            values = [str(headers[index])] + [str(row[index]) if index < len(row) else "" for row in rows]
            weights.append(max(4, min(40, max(map(len, values), default=4))))
        total = sum(weights)
        return [self.profile.content_width_cm() * weight / total for weight in weights]

    def add_table_part(self, 
        document: Document,
        headers: list[Any],
        rows: list[list[Any]],
        widths: list[float],
        alignments: list[str] | None = None,
        font_size: float = 14,
        font_family: str = "Times New Roman",
        line_spacing: float = 1.5,
        no_wrap_columns: set[int] | None = None,
        keep_together: bool = False,
    ):
        table = document.add_table(rows=1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.style = "Table Grid"
        table_width = sum(map(float, widths))
        tbl_pr = table._tbl.tblPr
        tbl_w = tbl_pr.find(qn("w:tblW"))
        if tbl_w is None:
            tbl_w = OxmlElement("w:tblW")
            tbl_pr.append(tbl_w)
        tbl_w.set(qn("w:type"), "dxa")
        tbl_w.set(qn("w:w"), str(round(table_width * 567)))

        grid = table._tbl.tblGrid
        for child in list(grid):
            grid.remove(child)
        for width in widths:
            col = OxmlElement("w:gridCol")
            col.set(qn("w:w"), str(round(float(width) * 567)))
            grid.append(col)

        all_rows = [headers] + rows
        for row_index, values in enumerate(all_rows):
            row = table.rows[0] if row_index == 0 else table.add_row()
            set_row_cant_split(row)
            if row_index == 0:
                set_repeat_table_header(row)
            for index, (cell, value, width) in enumerate(zip(row.cells, values, widths)):
                compact_cell = row_index > 0 and bool(no_wrap_columns) and index in no_wrap_columns
                set_cell_width(cell, float(width))
                if compact_cell:
                    set_cell_margins(cell, top=40, start=40, bottom=40, end=40)
                else:
                    set_cell_margins(cell)
                if no_wrap_columns and index in no_wrap_columns:
                    set_cell_no_wrap(cell)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                paragraph = cell.paragraphs[0]
                if row_index == 0:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.keep_with_next = True
                elif alignments:
                    value_text = str(value).strip().replace("−", "-")
                    if re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", value_text):
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        paragraph.alignment = {
                            "left": WD_ALIGN_PARAGRAPH.LEFT,
                            "center": WD_ALIGN_PARAGRAPH.CENTER,
                            "right": WD_ALIGN_PARAGRAPH.RIGHT,
                        }[alignments[index]]
                else:
                    value_text = str(value).strip().replace("−", "-")
                    paragraph.alignment = (
                        WD_ALIGN_PARAGRAPH.CENTER
                        if re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", value_text)
                        else WD_ALIGN_PARAGRAPH.LEFT
                    )
                paragraph.paragraph_format.first_line_indent = Cm(0)
                paragraph.paragraph_format.line_spacing = line_spacing
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                run = paragraph.add_run(clean_text(value, field=f"table cell {row_index},{index}"))
                run_font = font_family
                run_size = font_size
                # GOST 7.32-2017 6.1.1 permits bold only for section and structural
                # headings; tr23-26 3.7.13 does not request bold in tables, so the
                # header row stays regular.
                set_run_font(run, run_font, run_size)

        if keep_together and len(table.rows) > 1:
            # Word and LibreOffice honour keep-with-next inside table cells. Applying
            # it to every row except the last keeps a short logical table intact;
            # the option must not be used for a table taller than one page.
            for row in table.rows[:-1]:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.paragraph_format.keep_with_next = True

        return table

    def add_table(self, document: Document, block: dict[str, Any]) -> None:
        headers = block.get("headers")
        rows = block.get("rows")
        if not isinstance(headers, list) or not headers:
            raise SpecError("table requires non-empty headers")
        if not isinstance(rows, list):
            raise SpecError("table rows must be a list")
        if any(not isinstance(row, list) or len(row) != len(headers) for row in rows):
            raise SpecError("every table row must have the same number of cells as headers")
        widths = [float(value) for value in (block.get("widths_cm") or self.choose_widths(headers, rows))]
        if len(widths) != len(headers) or sum(widths) > self.profile.content_width_cm() + 0.01:
            raise SpecError(f"table widths must match columns and total at most {self.profile.content_width_cm()} cm")
        for header, width in zip(headers, widths):
            if str(header).strip() == "Итер." and width < 1.8:
                raise SpecError("the `Итер.` column must be at least 1.8 cm at Times New Roman 14 pt")
        alignments = block.get("column_alignments")
        if alignments is not None:
            if (
                not isinstance(alignments, list)
                or len(alignments) != len(headers)
                or any(value not in {"left", "center", "right"} for value in alignments)
            ):
                raise SpecError("column_alignments must match columns and use left, center, or right")
        long_token_mode = bool(block.get("long_token_mode", False))
        expected_font_size = float(self.profile.get("table_font_size"))
        expected_line_spacing = float(self.profile.get("table_line_spacing"))
        font_size = float(block.get("font_size", expected_font_size))
        font_family = str(block.get("font_family", "Times New Roman"))
        line_spacing = float(block.get("line_spacing", expected_line_spacing))
        if font_size != expected_font_size or font_family != "Times New Roman" or line_spacing != expected_line_spacing:
            raise SpecError(
                f"profile {self.profile.name} requires report tables to use "
                f"Times New Roman {expected_font_size:g} pt with {expected_line_spacing:g} line spacing"
            )
        no_wrap_raw = block.get("no_wrap_columns", [])
        if not isinstance(no_wrap_raw, list) or any(not isinstance(value, int) for value in no_wrap_raw):
            raise SpecError("no_wrap_columns must be a list of zero-based integer column indexes")
        no_wrap_columns = set(no_wrap_raw)
        if any(value < 0 or value >= len(headers) for value in no_wrap_columns):
            raise SpecError("no_wrap_columns contains an index outside the table")
        if no_wrap_columns or long_token_mode:
            raise SpecError(
                "long_token_mode and no_wrap_columns are disabled: preserve portrait pages and "
                "allow exact values to wrap naturally inside Word table cells"
            )

        number = clean_text(block.get("number", ""), field="table number")
        if not number:
            raise SpecError("table requires a number")
        caption = self.ctx.markup.normalize_report_text(clean_text(block.get("caption", ""), field="table caption"))
        if not caption:
            raise SpecError("table requires a caption")
        if caption.endswith("."):
            raise SpecError("table caption must not end with a period")

        rows_per_page_raw = block.get("rows_per_page")
        if rows_per_page_raw not in (None, "") and not self.profile.get("manual_table_continuations"):
            raise SpecError(
                "rows_per_page is disabled: keep one native Word table and use repeated headers"
            )
        if rows_per_page_raw not in (None, ""):
            rows_per_page = int(rows_per_page_raw)
            if rows_per_page < 1:
                raise SpecError("rows_per_page must be a positive integer")
            parts = [rows[index:index + rows_per_page] for index in range(0, len(rows), rows_per_page)] or [[]]
        else:
            parts = [rows]

        keep_together_raw = block.get("keep_together")
        keep_together = bool(keep_together_raw)
        # A short table that fits one page may stay whole: at 12 pt with single
        # spacing a ≤ 20-row table occupies roughly 10 cm or less, so moving it to
        # the next page keeps the previous page more than 60 % full. Longer tables
        # must use automatic pagination with repeated headers or rows_per_page.
        if keep_together and len(rows) > 20:
            raise SpecError(
                "keep_together is allowed only for tables that fit one page "
                "(≤ 20 rows); longer tables use rows_per_page or native pagination"
            )

        use_final_label = bool(block.get("use_final_label", False))
        if use_final_label:
            raise SpecError(
                "use_final_label is disabled in every profile: student_default uses native Word "
                "pagination without manual continuations, and SPbPU tr23-26 requires "
                "'Продолжение таблицы N' above every continuation part, including the last one"
            )

        keep_previous_reference(document)
        for part_index, part_rows in enumerate(parts):
            if part_index == 0:
                label = f"Таблица {number} – {caption}"
            elif use_final_label and part_index == len(parts) - 1:
                label = f"Окончание таблицы {number}"
            else:
                label = f"Продолжение таблицы {number}"
            caption_paragraph = document.add_paragraph(label, style="Report Table Caption")
            if part_index == 0:
                add_bookmark(document, caption_paragraph, object_bookmark_name("Table", number))
            # tr23-26 3.7.3: one free line between the text and the table; the
            # continuation captions sit at the top of a fresh page and need no gap.
            caption_paragraph.paragraph_format.space_before = Pt(OBJECT_GAP_PT if part_index == 0 else 0)
            caption_paragraph.paragraph_format.space_after = Pt(0)
            self.add_table_part(
                document,
                headers,
                part_rows,
                widths,
                alignments,
                font_size,
                font_family,
                line_spacing,
                no_wrap_columns,
                keep_together=keep_together and len(parts) == 1,
            )
            if part_index < len(parts) - 1:
                document.add_page_break()
        self.state.request_gap()
