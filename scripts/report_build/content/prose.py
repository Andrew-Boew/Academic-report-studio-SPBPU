"""Стратегии текстовых блоков."""

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
from report_build.document.markup import set_run_font
from report_build.spec.textutil import clean_text, keep_paragraph_lines

class ProseBuilder(ReportPart):
    """Заголовки, абзацы, списки и примечания."""

    def add_heading(self, document: Document, block: dict[str, Any]) -> None:
        self.state.clear_gap()
        level = int(block.get("level", 1))
        if level not in range(1, 7):
            raise SpecError(f"Heading level must be 1..6: {level}")
        text = self.ctx.markup.normalize_report_text(clean_text(block.get("text"), field="heading"))
        if MANUAL_HEADING_NUMBER.match(text):
            raise SpecError(
                f"Heading text must not contain a manually typed number; Word numbers headings automatically: {text!r}"
            )
        if "structural" in block:
            structural = bool(block["structural"])
        else:
            structural = level == 1 and bool(re.match(
                r"^(?:СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ВЫВОДЫ|СПИСОК|ПРИЛОЖЕНИЕ|РЕФЕРАТ|ABSTRACT)$", text, re.I
            ))
        style = "Report Structural Heading" if structural else f"Heading {level}"
        if structural and self.profile.get("uppercase_structural"):
            text = text.upper()
        paragraph = document.add_paragraph(text, style=style)
        if "page_break_before" in block:
            page_break_before = bool(block["page_break_before"])
        else:
            # GOST 7.32-2017 6.2.1 / tr23-26 3.3.1: every structural element and
            # every main-part section starts on a new page.
            page_break_before = structural or level == 1
        paragraph.paragraph_format.page_break_before = page_break_before
        if structural:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.space_before = Pt(0)
        else:
            # Match the confirmed Word layout exactly. The direct first-line
            # position overrides application-specific list defaults, so the
            # number begins at 1.25 cm in both Word and LibreOffice.
            paragraph.paragraph_format.left_indent = Cm(0)
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            paragraph.paragraph_format.line_spacing = 1.0
            if level == 1:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(18)
            else:
                paragraph.paragraph_format.space_before = Pt(12)
                paragraph.paragraph_format.space_after = Pt(6)

    def add_paragraph_block(self, document: Document, block: dict[str, Any]) -> None:
        paragraph = document.add_paragraph()
        self.state.apply_gap(paragraph.paragraph_format)
        alignments = {
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
        }
        alignment = str(block.get("align", "")).strip().lower()
        if alignment:
            if alignment not in alignments:
                raise SpecError(f"Unsupported paragraph alignment: {alignment}")
            if alignment == "left":
                raise SpecError(
                    "Body prose must stay justified (по ширине); do not request left alignment. "
                    "Rewrite or split paragraphs with long identifiers into shorter factual sentences instead."
                )
            paragraph.alignment = alignments[alignment]
        if "first_indent_cm" in block:
            paragraph.paragraph_format.first_line_indent = Cm(float(block["first_indent_cm"]))
        prefix = str(block.get("bold_prefix", "")).replace("—", "–")
        if PLACEHOLDER_PATTERN.search(prefix):
            raise SpecError(f"Placeholder detected in bold_prefix: {prefix[:80]}")
        if prefix and not prefix.endswith((" ", "\u00a0")):
            prefix += " "
        text = clean_text(block.get("text", ""), field="paragraph")
        if prefix:
            run = paragraph.add_run(prefix)
            set_run_font(run, "Times New Roman", 14, bold=True)
        if text:
            self.ctx.markup.add_report_text(paragraph, text)
        keep_paragraph_lines(paragraph)

    def add_note(self, document: Document, block: dict[str, Any]) -> None:
        """Add one or several GOST-style notes without imitating Word footnotes."""
        items = block.get("items")
        if items is None:
            text = self.ctx.markup.normalize_report_text(clean_text(block.get("text", ""), field="note"))
            if not text:
                raise SpecError("note requires text or a non-empty items list")
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            prefix = paragraph.add_run("Примечание – ")
            set_run_font(prefix, "Times New Roman", 14)
            self.ctx.markup.add_report_text(paragraph, text)
            keep_paragraph_lines(paragraph)
            return
        if not isinstance(items, list) or not items:
            raise SpecError("note items must be a non-empty list")
        heading = document.add_paragraph("Примечания")
        heading.paragraph_format.first_line_indent = Cm(1.25)
        for run in heading.runs:
            set_run_font(run, "Times New Roman", 14)
        for index, item in enumerate(items, 1):
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.left_indent = Cm(1.25)
            paragraph.paragraph_format.first_line_indent = Cm(-1.25)
            set_run_font(paragraph.add_run(f"{index} "), "Times New Roman", 14)
            self.ctx.markup.add_report_text(paragraph, clean_text(item, field=f"note item {index}"))
            keep_paragraph_lines(paragraph)

    def add_list(self, document: Document, block: dict[str, Any], numbered: bool) -> None:
        style = "List Number" if numbered else "List Bullet"
        items = block.get("items")
        if not isinstance(items, list) or not items:
            raise SpecError(f"{block.get('type')} requires a non-empty items list")
        for index, item in enumerate(items):
            if isinstance(item, dict):
                text = item.get("text", "")
                level = int(item.get("level", 1))
            else:
                text = item
                level = 1
            if level not in (1, 2, 3):
                raise SpecError("list item level must be 1, 2, or 3")
            paragraph = document.add_paragraph(style=style)
            if numbered and level > 1:
                ppr = paragraph._p.get_or_add_pPr()
                num_pr = ppr.find(qn("w:numPr"))
                if num_pr is None:
                    num_pr = OxmlElement("w:numPr")
                    ppr.append(num_pr)
                ilvl = num_pr.find(qn("w:ilvl"))
                if ilvl is None:
                    ilvl = OxmlElement("w:ilvl")
                    num_pr.insert(0, ilvl)
                ilvl.set(qn("w:val"), str(level - 1))
                style_ppr = document.styles[style]._element.find(qn("w:pPr"))
                style_num_pr = style_ppr.find(qn("w:numPr")) if style_ppr is not None else None
                style_num = style_num_pr.find(qn("w:numId")) if style_num_pr is not None else None
                num_id = num_pr.find(qn("w:numId"))
                if num_id is None and style_num is not None:
                    num_id = OxmlElement("w:numId")
                    num_id.set(qn("w:val"), style_num.get(qn("w:val")))
                    num_pr.append(num_id)
                paragraph.paragraph_format.left_indent = Cm((1066 + (level - 1) * 360) / 567)
                paragraph.paragraph_format.first_line_indent = Cm(-357 / 567)
            if index == 0:
                self.state.apply_gap(paragraph.paragraph_format)
            self.ctx.markup.add_report_text(paragraph, clean_text(text, field="list item"))
            keep_paragraph_lines(paragraph)
