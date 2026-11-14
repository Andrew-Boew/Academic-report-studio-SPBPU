"""Стратегия листинга."""

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
from report_build.document.cells import set_cell_margins, set_cell_width, set_row_cant_split
from report_build.document.markup import add_bookmark, object_bookmark_name, set_run_font
from report_build.spec.textutil import clean_text, keep_previous_reference

class ListingBuilder(ReportPart):
    """Короткие фрагменты кода и полные листинги приложений."""

    def add_code(self, document: Document, block: dict[str, Any]) -> None:
        source = str(block.get("source", "")).strip()
        if source:
            path = Path(source).expanduser().resolve()
            if not path.is_file():
                raise SpecError(f"Code source does not exist: {path}")
            text = path.read_text(encoding=str(block.get("encoding", "utf-8")), errors="replace")
        else:
            text = str(block.get("text", ""))
        if PLACEHOLDER_PATTERN.search(text):
            raise SpecError("Placeholder detected in code block")
        presentation = str(block.get("presentation", "fragment")).strip().lower()
        if presentation not in {"fragment", "appendix"}:
            raise SpecError("code presentation must be 'fragment' or 'appendix'")
        caption = str(block.get("caption", "")).strip().replace("—", "–")
        number = str(block.get("number", "")).strip()
        if presentation == "fragment":
            if not caption or not number:
                raise SpecError("framed code fragment requires caption and number")
            if caption.endswith("."):
                raise SpecError("listing caption must not end with a period")
            keep_previous_reference(document)
            paragraph = document.add_paragraph(f"Листинг {number} – {caption}", style="Report Table Caption")
            add_bookmark(document, paragraph, object_bookmark_name("Listing", number))
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(self.state.consume_before(TABLE_CAPTION_BEFORE_PT))
            paragraph.paragraph_format.space_after = Pt(TABLE_CAPTION_AFTER_PT)
            paragraph.paragraph_format.keep_with_next = True
        lines = text.expandtabs(4).splitlines() or [""]
        if presentation == "fragment":
            font_size = float(block.get("font_size", 12))
            if font_size not in {11.0, 12.0}:
                raise SpecError("framed code fragment font_size must be 11 or 12 pt")
            table = document.add_table(rows=1, cols=1)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False
            table.style = "Table Grid"
            tbl_pr = table._tbl.tblPr
            descriptor = OxmlElement("w:tblCaption")
            descriptor.set(qn("w:val"), "code-fragment")
            tbl_pr.append(descriptor)
            set_row_cant_split(table.rows[0])
            cell = table.cell(0, 0)
            set_cell_width(cell, self.profile.content_width_cm())
            set_cell_margins(cell, top=70, start=100, bottom=70, end=100)
            for index, line in enumerate(lines):
                paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
                paragraph.style = document.styles["Report Code Fragment"]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.first_line_indent = Cm(0)
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                set_run_font(paragraph.add_run(line), "Courier New", font_size)
            document.add_paragraph().paragraph_format.space_after = Pt(0)
        else:
            label = clean_text(block.get("label", ""), field="appendix code label")
            if not label:
                raise SpecError("appendix code requires a visible relative-path label")
            if Path(label).is_absolute() or ".." in Path(label).parts:
                raise SpecError("appendix code label must be a safe relative path")
            label_paragraph = document.add_paragraph(style="Report Appendix File Label")
            label_run = label_paragraph.add_run(f"Файл: {label}")
            set_run_font(label_run, "Times New Roman", 12, bold=True)
            paragraph = document.add_paragraph(style="Report Appendix Code")
            line_spacing_pt = block.get("line_spacing_pt")
            if line_spacing_pt is None:
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                paragraph.paragraph_format.line_spacing = 1.0
            else:
                line_spacing_pt = float(line_spacing_pt)
                if line_spacing_pt < 10.0 or line_spacing_pt > 12.0:
                    raise SpecError("appendix code line_spacing_pt must be within 10..12 pt")
                paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                paragraph.paragraph_format.line_spacing = Pt(line_spacing_pt)
            paragraph.paragraph_format.widow_control = True
            run = paragraph.add_run()
            for index, line in enumerate(lines):
                if index:
                    run.add_break()
                run.add_text(line)
            set_run_font(run, "Courier New", 10)
