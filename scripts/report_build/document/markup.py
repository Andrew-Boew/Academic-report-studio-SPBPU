"""Шрифт, гиперссылки и текст абзаца."""

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
from report_build.document.omath import latex_to_omath, style_math_runs

def set_run_font(run, name: str, size: float, *, bold: bool | None = None, italic: bool | None = None) -> None:
    run.font.name = name
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.bold = bold
    run.font.italic = italic
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{key}"), name)
    for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(f"w:{key}"), None)

def set_character_spacing(run, points: float) -> None:
    """Set expanded character spacing while keeping searchable text intact."""
    spacing = run._element.get_or_add_rPr().find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        run._element.get_or_add_rPr().append(spacing)
    spacing.set(qn("w:val"), str(round(points * 20)))

def add_external_hyperlink(paragraph, text: str, url: str) -> None:
    """Add a clickable URL that remains black and un-underlined per the local profile."""
    relation_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relation_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{key}"), "Times New Roman")
    rpr.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "28")
    rpr.append(size)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "none")
    rpr.append(underline)
    run.append(rpr)
    node = OxmlElement("w:t")
    node.text = text
    run.append(node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

def add_internal_hyperlink(paragraph, text: str, anchor: str) -> None:
    """Add a black, un-underlined link to a bookmark in the same document."""
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(f"w:{key}"), "Times New Roman")
    rpr.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "28")
    rpr.append(size)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "none")
    rpr.append(underline)
    run.append(rpr)
    node = OxmlElement("w:t")
    node.text = text
    run.append(node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

def object_bookmark_name(kind: str, number: str) -> str:
    """Return a Word-safe stable bookmark name for a numbered object."""
    normalized = number.strip().upper()
    match = re.fullmatch(r"([А-ЯЁ])\.(\d+)", normalized)
    if match:
        letter = CYRILLIC_BOOKMARK_LETTERS.get(match.group(1), f"CYR{ord(match.group(1))}")
        return f"_Report{kind}_{letter}_{match.group(2)}"
    return f"_Report{kind}_" + re.sub(r"[^A-Z0-9_]", "_", normalized)

def figure_bookmark_name(number: str) -> str:
    """Return a Word-safe stable bookmark name for a figure number."""
    return object_bookmark_name("Figure", number)

def add_bookmark(document: Document, paragraph, name: str) -> None:
    """Mark a caption as an internal-link destination without visible text."""
    used = [
        int(node.get(qn("w:id"), "0"))
        for node in document.element.xpath(".//w:bookmarkStart")
        if node.get(qn("w:id"), "").isdigit()
    ]
    bookmark_id = max(used, default=0) + 1
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(1 if paragraph._p.pPr is not None else 0, start)
    paragraph._p.append(end)

class MarkupWriter(ReportPart):
    """Строковый текст, сноски и внутренние ссылки абзаца."""

    def normalize_report_text(self, text: str) -> str:
        text = text.replace("—", "–")
        if self.profile.get("strict_units"):
            text = UNIT_PATTERN.sub("\u00a0", text)
            text = DEGREE_SPACE_PATTERN.sub("", text)
        return text

    def add_inline_math(self, paragraph, latex: str) -> None:
        """Insert a real inline OMML equation into the run flow at this position."""
        latex = latex.strip()
        if not latex:
            raise SpecError("inline math requires non-empty LaTeX")
        if re.search(r"\b(?:TODO|TBD|FIXME)\b|\[вставить[^\]]*\]", latex, re.I):
            raise SpecError(f"Placeholder detected in inline math: {latex[:80]}")
        if INLINE_MATH_PATTERN.search(latex):
            raise SpecError("nested inline math markers are not allowed")
        if re.search(r"\\\|\\begin\{", latex):
            raise SpecError(
                "inline math must stay short; use a separate equation block for multiline expressions"
            )
        equation = copy.deepcopy(latex_to_omath(latex))
        style_math_runs(equation)
        paragraph._p.append(equation)

    def add_footnote_reference(self, paragraph, note_text: str) -> None:
        """Attach a real Microsoft Word footnote to the current position."""
        self.state.footnotes.append(self.normalize_report_text(note_text.strip()))
        footnote_id = len(self.state.footnotes) + 1  # ids -1 and 0 are reserved separators
        run = paragraph.add_run()
        rpr = run._element.get_or_add_rPr()
        vert_align = OxmlElement("w:vertAlign")
        vert_align.set(qn("w:val"), "superscript")
        rpr.append(vert_align)
        reference = OxmlElement("w:footnoteReference")
        reference.set(qn("w:id"), str(footnote_id))
        run._r.append(reference)
        set_run_font(run, "Times New Roman", 14)

    def attach_footnotes(self, document: Document) -> None:
        """Create the word/footnotes.xml part with every collected footnote."""
        if not self.state.footnotes:
            return
        namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        root = etree.Element(f"{{{namespace}}}footnotes", nsmap={"w": namespace})
        for footnote_type, footnote_id, marker in (
            ("separator", "-1", "separator"),
            ("continuationSeparator", "0", "continuationSeparator"),
        ):
            footnote = etree.SubElement(root, f"{{{namespace}}}footnote")
            footnote.set(f"{{{namespace}}}type", footnote_type)
            footnote.set(f"{{{namespace}}}id", footnote_id)
            paragraph = etree.SubElement(footnote, f"{{{namespace}}}p")
            ppr = etree.SubElement(paragraph, f"{{{namespace}}}pPr")
            spacing = etree.SubElement(ppr, f"{{{namespace}}}spacing")
            spacing.set(f"{{{namespace}}}after", "0")
            spacing.set(f"{{{namespace}}}line", "240")
            spacing.set(f"{{{namespace}}}lineRule", "auto")
            run = etree.SubElement(paragraph, f"{{{namespace}}}r")
            etree.SubElement(run, f"{{{namespace}}}{marker}")
        for index, note_text in enumerate(self.state.footnotes, start=2):
            footnote = etree.SubElement(root, f"{{{namespace}}}footnote")
            footnote.set(f"{{{namespace}}}id", str(index))
            paragraph = etree.SubElement(footnote, f"{{{namespace}}}p")
            ppr = etree.SubElement(paragraph, f"{{{namespace}}}pPr")
            spacing = etree.SubElement(ppr, f"{{{namespace}}}spacing")
            spacing.set(f"{{{namespace}}}after", "0")
            spacing.set(f"{{{namespace}}}line", "240")
            spacing.set(f"{{{namespace}}}lineRule", "auto")
            indent = etree.SubElement(ppr, f"{{{namespace}}}ind")
            indent.set(f"{{{namespace}}}firstLine", "0")
            marker_run = etree.SubElement(paragraph, f"{{{namespace}}}r")
            marker_rpr = etree.SubElement(marker_run, f"{{{namespace}}}rPr")
            marker_vert = etree.SubElement(marker_rpr, f"{{{namespace}}}vertAlign")
            marker_vert.set(f"{{{namespace}}}val", "superscript")
            etree.SubElement(marker_run, f"{{{namespace}}}footnoteRef")
            text_run = etree.SubElement(paragraph, f"{{{namespace}}}r")
            text_rpr = etree.SubElement(text_run, f"{{{namespace}}}rPr")
            fonts = etree.SubElement(text_rpr, f"{{{namespace}}}rFonts")
            for key in ("ascii", "hAnsi", "eastAsia", "cs"):
                fonts.set(f"{{{namespace}}}{key}", "Times New Roman")
            size = etree.SubElement(text_rpr, f"{{{namespace}}}sz")
            size.set(f"{{{namespace}}}val", "20")
            node = etree.SubElement(text_run, f"{{{namespace}}}t")
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            node.text = f" {note_text}"
        content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
        relationship_type = getattr(
            RT, "FOOTNOTES", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
        )
        from docx.opc.packuri import PackURI
        from docx.opc.part import Part

        part = Part(
            PackURI("/word/footnotes.xml"),
            content_type,
            etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
            document.part.package,
        )
        document.part.relate_to(part, relationship_type)

    def _add_linked_text(self, paragraph, text: str) -> None:
        """Write one text segment; source, figure, table, listing and equation
        references become black internal hyperlinks to the object bookmarks."""
        matches: list[tuple[int, int, str, Any]] = []
        matches.extend((match.start(), match.end(), "source", match) for match in CITATION_PATTERN.finditer(text))
        matches.extend((match.start(), match.end(), "figure", match) for match in FIGURE_REFERENCE_PATTERN.finditer(text))
        matches.extend((match.start(), match.end(), "table", match) for match in TABLE_REFERENCE_PATTERN.finditer(text))
        matches.extend((match.start(), match.end(), "listing", match) for match in LISTING_REFERENCE_PATTERN.finditer(text))
        matches.extend((match.start(), match.end(), "equation", match) for match in EQUATION_REFERENCE_PATTERN.finditer(text))
        matches.sort(key=lambda item: item[0])
        cursor = 0
        for start, end, kind, match in matches:
            if start < cursor:
                continue
            if start > cursor:
                run = paragraph.add_run(text[cursor:start])
                set_run_font(run, "Times New Roman", 14)
            if kind == "source":
                number = int(match.group(1))
                add_internal_hyperlink(paragraph, match.group(0), f"_ReportBib{number}")
            elif kind == "figure":
                add_internal_hyperlink(paragraph, match.group(0), figure_bookmark_name(match.group("number")))
            elif kind == "table":
                add_internal_hyperlink(paragraph, match.group(0), object_bookmark_name("Table", match.group(1)))
            elif kind == "listing":
                add_internal_hyperlink(paragraph, match.group(0), object_bookmark_name("Listing", match.group(1)))
            else:
                number = match.group(1)
                if number in self.state.equation_numbers:
                    add_internal_hyperlink(paragraph, match.group(0), object_bookmark_name("Equation", number))
                else:
                    run = paragraph.add_run(match.group(0))
                    set_run_font(run, "Times New Roman", 14)
            cursor = end
        if cursor < len(text):
            run = paragraph.add_run(text[cursor:])
            set_run_font(run, "Times New Roman", 14)

    def add_report_text(self, paragraph, text: str) -> None:
        """Write body text; `[^текст]` becomes a real footnote, `$latex$` a real inline OMML equation."""
        text = self.normalize_report_text(text)
        segments: list[tuple[str, str]] = []
        cursor = 0
        for match in INLINE_MATH_PATTERN.finditer(text):
            if match.start() > cursor:
                segments.append(("text", text[cursor:match.start()]))
            segments.append(("math", match.group(1)))
            cursor = match.end()
        if cursor < len(text):
            segments.append(("text", text[cursor:]))
        if not any(kind == "math" for kind, _ in segments):
            segments = [("text", text)]
        for kind, value in segments:
            if kind == "math":
                self.add_inline_math(paragraph, value)
                continue
            footnote_matches = list(FOOTNOTE_PATTERN.finditer(value))
            if not footnote_matches:
                self._add_linked_text(paragraph, value)
                continue
            inner_cursor = 0
            for match in footnote_matches:
                if match.start() > inner_cursor:
                    self._add_linked_text(paragraph, value[inner_cursor:match.start()])
                self.add_footnote_reference(paragraph, match.group(1))
                inner_cursor = match.end()
            if inner_cursor < len(value):
                self._add_linked_text(paragraph, value[inner_cursor:])
