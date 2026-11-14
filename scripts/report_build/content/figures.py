"""Стратегия рисунка."""

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
from report_build.document.markup import add_bookmark, figure_bookmark_name
from report_build.spec.textutil import clean_text, keep_paragraph_lines, keep_previous_reference

class FigureBuilder(ReportPart):
    """Рисунки и пустые рамки под снимок."""

    def add_figure(self, document: Document, block: dict[str, Any]) -> None:
        path = Path(clean_text(block.get("path"), field="figure path")).expanduser().resolve()
        if not path.is_file():
            raise SpecError(f"Figure does not exist: {path}")
        width = float(block.get("width_cm", 14.5))
        if width <= 0 or width > self.profile.content_width_cm():
            raise SpecError(f"Figure width must be within 0..{self.profile.content_width_cm()} cm")
        keep_previous_reference(document)
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        # Moderate separation between the text and the figure.
        paragraph.paragraph_format.space_before = Pt(OBJECT_GAP_PT)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.keep_with_next = True
        run = paragraph.add_run()
        shape = run.add_picture(str(path), width=Cm(width))
        alt = str(block.get("alt_text", block.get("caption", ""))).strip()
        if alt:
            doc_pr = shape._inline.docPr
            doc_pr.set("descr", alt)
        annotation = block.get("annotation")
        if annotation:
            annotation_lines = annotation if isinstance(annotation, list) else [annotation]
            for annotation_line in annotation_lines:
                annotation_paragraph = document.add_paragraph(style="Report Figure Annotation")
                annotation_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                annotation_paragraph.paragraph_format.first_line_indent = Cm(0)
                annotation_paragraph.paragraph_format.space_before = Pt(0)
                annotation_paragraph.paragraph_format.space_after = Pt(0)
                annotation_paragraph.paragraph_format.keep_with_next = True
                self.ctx.markup.add_report_text(annotation_paragraph, clean_text(annotation_line, field="figure annotation"))
                keep_paragraph_lines(annotation_paragraph)
        number = clean_text(block.get("number", ""), field="figure number")
        if not number:
            raise SpecError("figure requires a number")
        caption = self.ctx.markup.normalize_report_text(clean_text(block.get("caption", ""), field="figure caption"))
        if not caption:
            raise SpecError("figure requires a caption")
        if caption.endswith("."):
            raise SpecError("figure caption must not end with a period")
        label = f"Рисунок {number} – {caption}"
        caption_paragraph = document.add_paragraph(label, style="Report Caption")
        add_bookmark(document, caption_paragraph, figure_bookmark_name(number))
        self.state.request_gap()

    def add_figure_placeholder(self, document: Document, block: dict[str, Any]) -> None:
        number = clean_text(block.get("number", ""), field="figure placeholder number")
        caption = clean_text(block.get("caption", ""), field="figure placeholder caption").replace("—", "–")
        if not number or not caption:
            raise SpecError("figure_placeholder requires number and caption")
        width = float(block.get("width_cm", 15.5))
        height = float(block.get("height_cm", 6.0))
        if width <= 0 or width > self.profile.content_width_cm():
            raise SpecError(f"Figure placeholder width must be within 0..{self.profile.content_width_cm()} cm")
        if height < 3.0 or height > 18.0:
            raise SpecError("Figure placeholder height must be within 3..18 cm")
        keep_previous_reference(document)
        table = document.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.style = "Table Grid"
        tbl_pr = table._tbl.tblPr
        descriptor = OxmlElement("w:tblCaption")
        descriptor.set(qn("w:val"), "figure-placeholder")
        tbl_pr.append(descriptor)
        row = table.rows[0]
        row.height = Cm(height)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        set_row_cant_split(row)
        cell = table.cell(0, 0)
        set_cell_width(cell, width)
        set_cell_margins(cell, top=0, start=100, bottom=0, end=100)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        paragraph = cell.paragraphs[0]
        paragraph.style = document.styles["Report Figure Placeholder"]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        # Intentionally keep the reserved area empty. Editorial instructions such as
        # "insert a screenshot" must never leak into the submitted report.
        label = f"Рисунок {number} – {caption}"
        caption_paragraph = document.add_paragraph(label, style="Report Caption")
        add_bookmark(document, caption_paragraph, figure_bookmark_name(number))
        self.state.request_gap()
