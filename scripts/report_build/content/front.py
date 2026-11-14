"""Стратегии элементов до и после основного текста."""

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
from report_build.document.markup import add_external_hyperlink, set_run_font
from report_build.document.styles import add_field
from report_build.spec.textutil import clean_text, count_spec_objects, keep_paragraph_lines

def structural_front_heading(document: Document, title: str) -> None:
    """A front-matter heading: uppercase, centred, own page, outside the TOC."""
    heading = document.add_paragraph(title.upper(), style="Report Structural Heading")
    heading.paragraph_format.page_break_before = True
    heading.paragraph_format.first_line_indent = Cm(0)
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), "9")
    heading._p.get_or_add_pPr().append(outline)

class FrontMatterBuilder(ReportPart):
    """Содержание, реферат, задание, перечень, источники и приложения."""

    def add_toc(self, document: Document, levels: int) -> None:
        title = "СОДЕРЖАНИЕ" if self.profile.get("uppercase_structural") else "Содержание"
        heading = document.add_paragraph(title, style="Report Structural Heading")
        heading.paragraph_format.page_break_before = True
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.paragraph_format.first_line_indent = Cm(0)
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), "9")
        heading._p.get_or_add_pPr().append(outline)
        paragraph = document.add_paragraph(style="Report TOC Field")
        paragraph.paragraph_format.first_line_indent = Cm(0)
        add_field(paragraph, f' TOC \\o "1-{levels}" \\h \\z \\u ', "Обновите содержание в Word")

    def add_abbreviations(self, document: Document, block: dict[str, Any]) -> None:
        items = block.get("items")
        if not isinstance(items, list) or not items:
            raise SpecError("abbreviations requires a non-empty items list")
        if len(items) <= 3:
            raise SpecError(
                "Перечень сокращений составляют только при более чем трёх обозначениях "
                "(tr23-26 2.4.2); при меньшем числе раскройте сокращение в тексте при первом упоминании"
            )
        parsed: list[tuple[str, str]] = []
        for index, item in enumerate(items):
            if not isinstance(item, dict):
                raise SpecError(f"abbreviations item {index} must be an object")
            term = clean_text(item.get("term"), field=f"abbreviations[{index}].term")
            definition = clean_text(item.get("definition"), field=f"abbreviations[{index}].definition")
            parsed.append((term, definition))
        parsed.sort(key=lambda pair: pair[0].casefold())
        self.state.clear_gap()
        title = "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ" if self.profile.get("uppercase_structural") else "Перечень сокращений и обозначений"
        heading = document.add_paragraph(title, style="Report Structural Heading")
        heading.paragraph_format.page_break_before = True
        heading.paragraph_format.first_line_indent = Cm(0)
        lead_in = document.add_paragraph()
        lead_in.paragraph_format.first_line_indent = Cm(1.25)
        self.ctx.markup.add_report_text(lead_in, "В настоящей работе применяют следующие сокращения и обозначения:")
        keep_paragraph_lines(lead_in)
        for term, definition in parsed:
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.left_indent = Cm(4.0)
            paragraph.paragraph_format.first_line_indent = Cm(-4.0)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(4.0), WD_TAB_ALIGNMENT.LEFT)
            set_run_font(paragraph.add_run(term), "Times New Roman", 14)
            set_run_font(paragraph.add_run("\t– "), "Times New Roman", 14)
            self.ctx.markup.add_report_text(paragraph, definition)
            keep_paragraph_lines(paragraph)

    def add_abstract(self, document: Document, block: dict[str, Any]) -> None:
        """РЕФЕРАТ and its English ABSTRACT (tr23-26 2.3)."""
        keywords = block.get("keywords")
        ru_text = block.get("text")
        en = block.get("en") if isinstance(block.get("en"), dict) else {}
        en_keywords = en.get("keywords")
        en_text = en.get("text")
        for label, value, low, high in (
            ("abstract.keywords", keywords, 5, 15),
            ("abstract.en.keywords", en_keywords, 5, 15),
        ):
            if not isinstance(value, list) or not (low <= len(value) <= high):
                raise SpecError(f"{label} must contain 5..15 key words")
        for label, value in (("abstract.text", ru_text), ("abstract.en.text", en_text)):
            if not isinstance(value, list) or not value:
                raise SpecError(f"{label} must be a non-empty list of paragraphs")
        figures, tables, sources, appendices = count_spec_objects(document._report_blocks)  # type: ignore[attr-defined]
        volume_line = (
            f"Работа содержит: __ с., {figures} рис., {tables} табл., {sources} источ., {appendices} прил."
        )
        self.state.clear_gap()
        structural_front_heading(document, "Реферат")
        volume_paragraph = document.add_paragraph()
        volume_paragraph.paragraph_format.first_line_indent = Cm(1.25)
        set_run_font(volume_paragraph.add_run(volume_line), "Times New Roman", 14)
        keep_paragraph_lines(volume_paragraph)
        keywords_paragraph = document.add_paragraph()
        keywords_paragraph.paragraph_format.first_line_indent = Cm(0)
        set_run_font(
            keywords_paragraph.add_run("КЛЮЧЕВЫЕ СЛОВА: " + ", ".join(str(word).upper() for word in keywords)),
            "Times New Roman", 14,
        )
        keep_paragraph_lines(keywords_paragraph)
        for item in ru_text:
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            self.ctx.markup.add_report_text(paragraph, clean_text(item, field="abstract text"))
            keep_paragraph_lines(paragraph)

        structural_front_heading(document, "Abstract")
        en_volume_paragraph = document.add_paragraph()
        en_volume_paragraph.paragraph_format.first_line_indent = Cm(1.25)
        set_run_font(
            en_volume_paragraph.add_run(
                f"The work contains: __ pages, {figures} figures, {tables} tables, {sources} sources, {appendices} appendices."
            ),
            "Times New Roman", 14,
        )
        keep_paragraph_lines(en_volume_paragraph)
        en_keywords_paragraph = document.add_paragraph()
        en_keywords_paragraph.paragraph_format.first_line_indent = Cm(0)
        set_run_font(
            en_keywords_paragraph.add_run("KEYWORDS: " + ", ".join(str(word).upper() for word in en_keywords)),
            "Times New Roman", 14,
        )
        keep_paragraph_lines(en_keywords_paragraph)
        for item in en_text:
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            self.ctx.markup.add_report_text(paragraph, clean_text(item, field="abstract en text"))
            keep_paragraph_lines(paragraph)

    def add_assignment(self, document: Document, block: dict[str, Any]) -> None:
        """ЗАДАНИЕ on the work (tr23-26 Приложение Б shape)."""
        fields = block.get("fields")
        if not isinstance(fields, list) or not fields:
            raise SpecError("assignment requires a non-empty fields list")
        self.state.clear_gap()
        structural_front_heading(document, "Задание")
        subtitle = document.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.paragraph_format.first_line_indent = Cm(0)
        set_run_font(subtitle.add_run("на выполнение выпускной квалификационной работы"), "Times New Roman", 14, bold=True)
        keep_paragraph_lines(subtitle)
        for index, field in enumerate(fields, 1):
            if not isinstance(field, dict):
                raise SpecError(f"assignment field {index} must be an object")
            label = clean_text(field.get("label"), field=f"assignment[{index}].label")
            value = str(field.get("value", "")).strip()
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.first_line_indent = Cm(1.25)
            set_run_font(paragraph.add_run(f"{index}. {label}: "), "Times New Roman", 14)
            if value:
                self.ctx.markup.add_report_text(paragraph, value)
            keep_paragraph_lines(paragraph)

    def add_appendix_heading(self, document: Document, block: dict[str, Any]) -> None:
        letter = clean_text(block.get("letter"), field="appendix letter").upper()
        title = clean_text(block.get("title", ""), field="appendix title")
        forbidden_letters = {"Ё", "З", "Й", "О", "Ч", "Ь", "Ы", "Ъ"}
        if self.profile.get("strict_appendices") and letter in forbidden_letters:
            raise SpecError(f"appendix letter {letter!r} is not permitted by profile {self.profile.name}")
        prefix = "ПРИЛОЖЕНИЕ" if self.profile.get("strict_appendices") else "Приложение"
        paragraph = document.add_paragraph(f"{prefix} {letter}", style="Report Structural Heading")
        paragraph.paragraph_format.page_break_before = True
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        if title:
            title_paragraph = document.add_paragraph(title, style="Normal")
            if self.profile.get("strict_appendices"):
                title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                title_paragraph.paragraph_format.first_line_indent = Cm(0)
                for run in title_paragraph.runs:
                    set_run_font(run, "Times New Roman", 14, bold=True)

    def add_bibliography(self, document: Document, block: dict[str, Any]) -> None:
        self.state.clear_gap()
        title = clean_text(block.get("title", "Список использованных источников"), field="bibliography title")
        if self.profile.get("uppercase_structural"):
            title = title.upper()
        heading = document.add_paragraph(title, style="Report Structural Heading")
        heading.paragraph_format.page_break_before = True
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.paragraph_format.first_line_indent = Cm(0)
        items = block.get("items")
        if not isinstance(items, list) or not items:
            raise SpecError("bibliography requires at least one verified source")
        for index, item in enumerate(items, 1):
            if not isinstance(item, dict):
                raise SpecError(
                    "Every bibliography item must be an object with kind, citation, url, accessed, and verified=true"
                )
            kind = str(item.get("kind", "")).strip()
            if kind not in SOURCE_KINDS:
                raise SpecError(f"Unsupported bibliography source kind: {kind!r}")
            if item.get("verified") is not True:
                raise SpecError("Every bibliography source must be opened and verified before assembly")
            citation = clean_text(item.get("citation"), field="bibliography citation").rstrip(". ")
            url = clean_text(item.get("url"), field="bibliography url").rstrip(".,;:)")
            parsed = urlparse(url)
            host = (parsed.hostname or "").casefold()
            if parsed.scheme != "https" or not host or host in {"localhost", "127.0.0.1", "::1"}:
                raise SpecError(f"Bibliography URL must be a public HTTPS link: {url!r}")
            accessed = clean_text(item.get("accessed"), field="bibliography access date")
            try:
                dt.datetime.strptime(accessed, "%d.%m.%Y")
            except ValueError as exc:
                raise SpecError(f"Bibliography accessed date must use DD.MM.YYYY: {accessed!r}") from exc
            paragraph = document.add_paragraph(style="Report Bibliography")
            bookmark_start = OxmlElement("w:bookmarkStart")
            bookmark_start.set(qn("w:id"), str(50000 + index))
            bookmark_start.set(qn("w:name"), f"_ReportBib{index}")
            paragraph._p.insert(1 if paragraph._p.pPr is not None else 0, bookmark_start)
            # SPbPU tr23-26 2.9.2 numbers sources with Arabic numerals without a
            # trailing period; the confirmed student profile keeps `1. `.
            source_prefix = f"{index} " if self.profile.get("uppercase_structural") else f"{index}. "
            prefix = paragraph.add_run(f"{source_prefix}{citation}. URL: ")
            set_run_font(prefix, "Times New Roman", 14)
            add_external_hyperlink(paragraph, url, url)
            suffix = paragraph.add_run(f" (дата обращения: {accessed}).")
            set_run_font(suffix, "Times New Roman", 14)
            bookmark_end = OxmlElement("w:bookmarkEnd")
            bookmark_end.set(qn("w:id"), str(50000 + index))
            paragraph._p.append(bookmark_end)
