"""Стили Word, списки и секции страницы."""

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

def configure_style(style, *, font="Times New Roman", size=14, bold=None, italic=None,
                    align=None, first_indent_cm=None, left_indent_cm=None,
                    line_spacing=1.5, before=0, after=0, keep_next=None) -> None:
    style.font.name = font
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    style.font.color.rgb = RGBColor(0, 0, 0)
    rfonts = style._element.get_or_add_rPr().rFonts
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{key}"), font)
    for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rfonts.attrib.pop(qn(f"w:{key}"), None)
    pf = style.paragraph_format
    pf.alignment = align
    pf.first_line_indent = Cm(first_indent_cm) if first_indent_cm is not None else None
    pf.left_indent = Cm(left_indent_cm) if left_indent_cm is not None else None
    pf.line_spacing = line_spacing
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = keep_next
    pf.widow_control = True

def ensure_style(document: Document, name: str, style_type=WD_STYLE_TYPE.PARAGRAPH):
    try:
        return document.styles[name]
    except KeyError:
        return document.styles.add_style(name, style_type)

def set_outline_level(style, level: int) -> None:
    ppr = style._element.get_or_add_pPr()
    node = ppr.find(qn("w:outlineLvl"))
    if node is None:
        node = OxmlElement("w:outlineLvl")
        ppr.append(node)
    node.set(qn("w:val"), str(level))

def set_style_numbering(style, *, level: int, num_id: int) -> None:
    ppr = style._element.get_or_add_pPr()
    old = ppr.find(qn("w:numPr"))
    if old is not None:
        ppr.remove(old)
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level))
    num_id_node = OxmlElement("w:numId")
    num_id_node.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl)
    num_pr.append(num_id_node)
    ppr.append(num_pr)

def configure_dash_bullet_numbering(document: Document) -> int:
    """Link List Bullet to a real Word list whose marker is a dash (tr23-26 3.5.2)."""
    numbering = document.part.numbering_part.element
    abstract_ids = [
        int(node.get(qn("w:abstractNumId"), "0"))
        for node in numbering.findall(qn("w:abstractNum"))
    ]
    num_ids = [
        int(node.get(qn("w:numId"), "0"))
        for node in numbering.findall(qn("w:num"))
    ]
    abstract_id = max(abstract_ids, default=-1) + 1
    num_id = max(num_ids, default=0) + 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), "bullet")
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), "–")
    suffix = OxmlElement("w:suff")
    suffix.set(qn("w:val"), "space")
    lvl_jc = OxmlElement("w:lvlJc")
    lvl_jc.set(qn("w:val"), "left")
    ppr = OxmlElement("w:pPr")
    indent = OxmlElement("w:ind")
    # The dash starts at the 1.25 cm paragraph indent; wrapped lines align
    # 0.63 cm further right, right after the dash and its space.
    indent.set(qn("w:left"), "1066")
    indent.set(qn("w:hanging"), "357")
    ppr.append(indent)
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{key}"), "Times New Roman")
    rpr.append(fonts)
    for child in (start, num_fmt, lvl_text, suffix, lvl_jc, ppr, rpr):
        lvl.append(child)
    abstract.append(lvl)

    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        numbering.insert(numbering.index(first_num), abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    reference = OxmlElement("w:abstractNumId")
    reference.set(qn("w:val"), str(abstract_id))
    num.append(reference)
    numbering.append(num)
    set_style_numbering(document.styles["List Bullet"], level=0, num_id=num_id)

    style = document.styles["List Bullet"]
    style.paragraph_format.left_indent = Cm(1066 / 567)
    style.paragraph_format.first_line_indent = Cm(-357 / 567)
    return num_id

def configure_paren_number_numbering(document: Document) -> int:
    """Link List Number to a real Word list labelled `1)`, `2)`, ... (tr23-26 3.5.3)."""
    numbering = document.part.numbering_part.element
    abstract_ids = [
        int(node.get(qn("w:abstractNumId"), "0"))
        for node in numbering.findall(qn("w:abstractNum"))
    ]
    num_ids = [
        int(node.get(qn("w:numId"), "0"))
        for node in numbering.findall(qn("w:num"))
    ]
    abstract_id = max(abstract_ids, default=-1) + 1
    num_id = max(num_ids, default=0) + 1

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    num_fmt = OxmlElement("w:numFmt")
    num_fmt.set(qn("w:val"), "decimal")
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), "%1)")
    suffix = OxmlElement("w:suff")
    suffix.set(qn("w:val"), "space")
    lvl_jc = OxmlElement("w:lvlJc")
    lvl_jc.set(qn("w:val"), "left")
    ppr = OxmlElement("w:pPr")
    indent = OxmlElement("w:ind")
    indent.set(qn("w:left"), "1066")
    indent.set(qn("w:hanging"), "357")
    ppr.append(indent)
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{key}"), "Times New Roman")
    rpr.append(fonts)
    for child in (start, num_fmt, lvl_text, suffix, lvl_jc, ppr, rpr):
        lvl.append(child)
    abstract.append(lvl)

    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        numbering.insert(numbering.index(first_num), abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    reference = OxmlElement("w:abstractNumId")
    reference.set(qn("w:val"), str(abstract_id))
    num.append(reference)
    numbering.append(num)
    set_style_numbering(document.styles["List Number"], level=0, num_id=num_id)
    return num_id

def add_field(paragraph, instruction: str, placeholder: str = "") -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend((begin, instr, separate, text, end))
    set_run_font(run, "Times New Roman", 14)

def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    add_field(paragraph, " PAGE ", "1")

def add_update_fields(document: Document) -> None:
    settings = document.settings._element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")

def add_centered_paragraph(document: Document, text: str = "", size: float = 14,
                           bold: bool = False, after: float = 0) -> Any:
    paragraph = document.add_paragraph(style="Report Cover")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_after = Pt(after)
    run = paragraph.add_run(text)
    set_run_font(run, "Times New Roman", size, bold=bold)
    return paragraph

class StyleBuilder(ReportPart):
    """Поля, стили и нумерация документа."""

    def configure_heading_numbering(self, document: Document) -> int:
        """Create a real Word multilevel list linked to Heading 1–6."""
        numbering = document.part.numbering_part.element
        abstract_ids = [
            int(node.get(qn("w:abstractNumId"), "0"))
            for node in numbering.findall(qn("w:abstractNum"))
        ]
        num_ids = [
            int(node.get(qn("w:numId"), "0"))
            for node in numbering.findall(qn("w:num"))
        ]
        abstract_id = max(abstract_ids, default=-1) + 1
        num_id = max(num_ids, default=0) + 1

        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))
        multi = OxmlElement("w:multiLevelType")
        multi.set(qn("w:val"), "multilevel")
        abstract.append(multi)
        for level, label in enumerate(self.profile.get("heading_labels")):
            lvl = OxmlElement("w:lvl")
            lvl.set(qn("w:ilvl"), str(level))
            start = OxmlElement("w:start")
            start.set(qn("w:val"), "1")
            num_fmt = OxmlElement("w:numFmt")
            num_fmt.set(qn("w:val"), "decimal")
            p_style = OxmlElement("w:pStyle")
            p_style.set(qn("w:val"), f"Heading{level + 1}")
            lvl_text = OxmlElement("w:lvlText")
            lvl_text.set(qn("w:val"), label)
            suffix = OxmlElement("w:suff")
            suffix.set(qn("w:val"), "space")
            justification = OxmlElement("w:lvlJc")
            justification.set(qn("w:val"), "left")
            ppr = OxmlElement("w:pPr")
            indent = OxmlElement("w:ind")
            # Number position = left - hanging = 709 twips (1.25 cm).
            # Heading text begins at 1417 twips (2.50 cm). Heading styles themselves
            # keep zero direct indents so the positions are never added twice.
            indent.set(qn("w:left"), "1417")
            indent.set(qn("w:hanging"), "708")
            ppr.append(indent)
            for child in (start, num_fmt, p_style, lvl_text, suffix, justification, ppr):
                lvl.append(child)
            abstract.append(lvl)

        first_num = numbering.find(qn("w:num"))
        if first_num is None:
            numbering.append(abstract)
        else:
            numbering.insert(numbering.index(first_num), abstract)
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        reference = OxmlElement("w:abstractNumId")
        reference.set(qn("w:val"), str(abstract_id))
        num.append(reference)
        numbering.append(num)

        labels = self.profile.get("heading_labels")
        for level in range(len(labels)):
            set_style_numbering(document.styles[f"Heading {level + 1}"], level=level, num_id=num_id)
        return num_id

    def configure_document(self, document: Document, page_start: int) -> None:
        section = document.sections[0]
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(float(self.profile.get("left_margin_cm")))
        section.right_margin = Cm(float(self.profile.get("right_margin_cm")))
        section.top_margin = Cm(float(self.profile.get("top_margin_cm")))
        section.bottom_margin = Cm(float(self.profile.get("bottom_margin_cm")))
        section.header_distance = Cm(1.0)
        section.footer_distance = Cm(1.0)
        section.different_first_page_header_footer = True

        sect_pr = section._sectPr
        pg_num_type = sect_pr.find(qn("w:pgNumType"))
        if pg_num_type is None:
            pg_num_type = OxmlElement("w:pgNumType")
            sect_pr.append(pg_num_type)
        pg_num_type.set(qn("w:start"), str(page_start))

        settings = document.settings._element
        auto_hyphenation = settings.find(qn("w:autoHyphenation"))
        if auto_hyphenation is None:
            auto_hyphenation = OxmlElement("w:autoHyphenation")
            settings.append(auto_hyphenation)
        auto_hyphenation.set(qn("w:val"), "true" if self.profile.get("automatic_hyphenation") else "false")

        normal = document.styles["Normal"]
        configure_style(
            normal,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_indent_cm=1.25,
            line_spacing=1.5,
        )

        configure_style(
            document.styles["Heading 1"], size=18, bold=True,
            align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0, left_indent_cm=0,
            line_spacing=1.0, before=0, after=18, keep_next=True,
        )
        configure_style(
            document.styles["Heading 2"], size=16, bold=True,
            align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0, left_indent_cm=0,
            line_spacing=1.0, before=12, after=6, keep_next=True,
        )
        for level in range(3, 6):
            configure_style(
                document.styles[f"Heading {level}"], size=14, bold=True, italic=False,
                align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0, left_indent_cm=0,
                line_spacing=1.0, before=12, after=6, keep_next=True,
            )
        configure_style(
            document.styles["Heading 6"], size=14, bold=False, italic=True,
            align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0, left_indent_cm=0,
            line_spacing=1.0, before=12, after=6, keep_next=True,
        )
        structural_heading = ensure_style(document, "Report Structural Heading")
        configure_style(
            structural_heading, size=16, bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
            line_spacing=1.5, before=0, after=0, keep_next=True,
        )
        set_outline_level(structural_heading, 0)
        self.configure_heading_numbering(document)

        for style_name in (
            "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6",
            "Report Structural Heading",
        ):
            ppr = document.styles[style_name]._element.get_or_add_pPr()
            suppress = ppr.find(qn("w:suppressAutoHyphens"))
            if suppress is None:
                suppress = OxmlElement("w:suppressAutoHyphens")
                ppr.append(suppress)
            suppress.set(qn("w:val"), "true")

        configure_style(
            ensure_style(document, "Report Caption"), size=14,
            align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
            line_spacing=float(self.profile.get("caption_line_spacing")), before=0, after=0, keep_next=False,
        )
        configure_style(
            ensure_style(document, "Report Table Caption"), size=14,
            align=(WD_ALIGN_PARAGRAPH.LEFT if self.profile.get("table_caption_alignment") == "left" else WD_ALIGN_PARAGRAPH.CENTER),
            first_indent_cm=0,
            line_spacing=float(self.profile.get("caption_line_spacing")), before=0, after=0, keep_next=True,
        )
        configure_style(
            ensure_style(document, "Report Figure Annotation"), size=14,
            align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0,
            line_spacing=1.5, before=0, after=0, keep_next=True,
        )
        configure_style(
            ensure_style(document, "Report Code Fragment"), font="Courier New", size=12,
            align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0,
            line_spacing=1.0, before=0, after=0,
        )
        configure_style(
            ensure_style(document, "Report Appendix Code"), font="Courier New", size=10,
            align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0,
            line_spacing=1.0, before=0, after=0,
        )
        configure_style(
            ensure_style(document, "Report Appendix File Label"), font="Times New Roman", size=12,
            bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0,
            line_spacing=1.0, before=6, after=0, keep_next=True,
        )
        configure_style(
            ensure_style(document, "Report Figure Placeholder"), size=12, italic=True,
            align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
            line_spacing=1.0, before=0, after=0, keep_next=True,
        )
        configure_style(
            ensure_style(document, "Report Equation"), font="Cambria Math", size=14,
            align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
            line_spacing=1.5, before=0, after=0,
        )
        for half_points in (28,):
            equation_style = ensure_style(document, f"Report Equation {half_points}")
            equation_style.base_style = document.styles["Report Equation"]
            configure_style(
                equation_style, font="Cambria Math", size=half_points / 2,
                align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
                line_spacing=1.5, before=0, after=0,
            )
        configure_style(
            ensure_style(document, "Report Cover"), size=14,
            align=WD_ALIGN_PARAGRAPH.CENTER, first_indent_cm=0,
            line_spacing=1.15, before=0, after=0,
        )
        configure_style(
            ensure_style(document, "Report TOC Field"), size=14,
            align=WD_ALIGN_PARAGRAPH.LEFT, first_indent_cm=0,
            line_spacing=1.5, before=0, after=0,
        )
        configure_style(
            ensure_style(document, "Report Bibliography"), size=14,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_indent_cm=-0.75,
            left_indent_cm=1.25, line_spacing=1.5,
        )

        for name, left, first in (
            ("List Bullet", 1066 / 567, -357 / 567),
            ("List Number", 1066 / 567, -357 / 567),
        ):
            style = document.styles[name]
            configure_style(
                style, size=14, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                first_indent_cm=first, left_indent_cm=left,
                line_spacing=1.5,
            )

        if self.profile.get("dash_bullet_lists"):
            configure_dash_bullet_numbering(document)
            configure_paren_number_numbering(document)

        toc_left_indents_cm = self.profile.get("toc_indents_cm")
        for level in range(1, 4):
            name = f"TOC {level}"
            configure_style(
                ensure_style(document, name), size=14,
                align=WD_ALIGN_PARAGRAPH.LEFT,
                first_indent_cm=0,
                left_indent_cm=toc_left_indents_cm[level],
                line_spacing=1.5,
                after=0,
            )

        # The first (front-matter) section keeps empty footers: the title page,
        # assignment, abstract and abbreviation list are counted but their numbers
        # are not printed (tr23-26 3.4.2). The body section gets the PAGE field.
        add_update_fields(document)

    def start_body_section(self, document: Document):
        """Close the front-matter section and open the numbered body section."""
        from docx.enum.section import WD_SECTION_START

        body_section = document.add_section(WD_SECTION_START.NEW_PAGE)
        sect_pr = body_section._sectPr
        page_numbering = sect_pr.find(qn("w:pgNumType"))
        if page_numbering is not None:
            sect_pr.remove(page_numbering)  # continue the section-1 counter
        body_section.different_first_page_header_footer = False
        body_section.footer.is_linked_to_previous = False
        add_page_number(body_section.footer.paragraphs[0])
        return body_section
