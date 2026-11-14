"""Стратегия выносной формулы."""

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
from report_build.document.markup import add_bookmark, object_bookmark_name, set_run_font
from report_build.document.omath import latex_to_omath, style_math_runs

class EquationBuilder(ReportPart):
    """Выносные формулы Word."""

    def add_equation(self, document: Document, block: dict[str, Any]) -> None:
        self.state.clear_gap()
        paragraph = document.add_paragraph(style="Report Equation")
        # Moderate separation between the text and the formula; the gap below is
        # requested from the next paragraph.
        paragraph.paragraph_format.space_before = Pt(OBJECT_GAP_PT)
        number = str(block.get("number", "")).strip()
        if number:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.tab_stops.add_tab_stop(
                Cm(self.profile.content_width_cm() / 2), WD_TAB_ALIGNMENT.CENTER
            )
            paragraph.paragraph_format.tab_stops.add_tab_stop(
                Cm(self.profile.content_width_cm()), WD_TAB_ALIGNMENT.RIGHT
            )
            paragraph.add_run().add_tab()
        latex = str(block.get("latex", "")).strip()
        if not latex:
            raise SpecError("equation requires non-empty LaTeX")
        if re.search(r"\b(?:TODO|TBD|FIXME)\b|\[вставить[^\]]*\]", latex, re.I):
            raise SpecError(f"Placeholder detected in equation latex: {latex[:80]}")
        if re.search(r"\\\\|\\begin\{(?:aligned|alignedat|array|matrix|pmatrix|bmatrix|vmatrix|Vmatrix)\}", latex):
            raise SpecError(
                "one equation block must contain one equation or parameter value; "
                "create separate equation blocks instead of aligned arrays or matrices"
            )
        longest_number = max((len(value) for value in re.findall(r"\d+", latex)), default=0)
        if longest_number > 60:
            raise SpecError(
                "an exact integer longer than 60 digits must be represented by named decimal "
                "blocks in separate 14 pt equations"
            )
        equation_half_points = "28"
        paragraph.style = document.styles[f"Report Equation {equation_half_points}"]
        equation = copy.deepcopy(latex_to_omath(latex))
        style_math_runs(equation, equation_half_points)
        paragraph_run_properties = paragraph._p.get_or_add_pPr().find(qn("w:rPr"))
        if paragraph_run_properties is None:
            paragraph_run_properties = OxmlElement("w:rPr")
            paragraph._p.get_or_add_pPr().append(paragraph_run_properties)
        paragraph_fonts = paragraph_run_properties.find(qn("w:rFonts"))
        if paragraph_fonts is None:
            paragraph_fonts = OxmlElement("w:rFonts")
            paragraph_run_properties.insert(0, paragraph_fonts)
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            paragraph_fonts.set(qn(f"w:{key}"), "Cambria Math")
        for tag in ("w:sz", "w:szCs"):
            size = paragraph_run_properties.find(qn(tag))
            if size is None:
                size = OxmlElement(tag)
                paragraph_run_properties.append(size)
            size.set(qn("w:val"), equation_half_points)
        paragraph._p.append(equation)
        if number:
            run = paragraph.add_run()
            run.add_tab()
            run.add_text(f"({number})")
            set_run_font(run, "Times New Roman", 14)
            add_bookmark(document, paragraph, object_bookmark_name("Equation", number))
        self.state.request_gap()
