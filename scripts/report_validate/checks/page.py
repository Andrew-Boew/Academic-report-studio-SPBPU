"""Страница, поля и колонтитул с номером."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class PageCheck:
    """Формат A4, поля и номер страницы."""

    def apply(self, report) -> None:
        if not report.document.sections:
            report.errors.append("В документе отсутствуют секции")
        for index, section in enumerate(report.document.sections, 1):
            prefix = f"Секция {index}"
            if section.orientation != WD_ORIENT.PORTRAIT:
                report.errors.append(f"{prefix}: требуется строго книжная ориентация страницы")
            expected_width = 21.0
            expected_height = 29.7
            if not near(section.page_width.cm, expected_width) or not near(section.page_height.cm, expected_height):
                report.errors.append(
                    f"{prefix}: ожидается формат A4, получено "
                    f"{section.page_width.cm:.2f} × {section.page_height.cm:.2f} см"
                )
            checks = (
                ("левое поле", section.left_margin.cm, 3.0),
                ("правое поле", section.right_margin.cm, report.expected_right_margin),
                ("верхнее поле", section.top_margin.cm, 2.0),
                ("нижнее поле", section.bottom_margin.cm, 2.0),
            )
            for label, actual, expected in checks:
                if not near(actual, expected):
                    report.errors.append(f"{prefix}: {label} {actual:.2f} см вместо {expected:.2f} см")

        first_section = report.document.sections[0]
        if not first_section.different_first_page_header_footer:
            report.errors.append("На титульном листе не отключён показ номера страницы")
        page_numbering = first_section._sectPr.find(qn("w:pgNumType"))
        page_start = page_numbering.get(qn("w:start")) if page_numbering is not None else None
        if page_start != report.expected_page_start:
            report.errors.append(
                f"Профиль {report.format_profile}: начальное значение счётчика страниц должно быть {report.expected_page_start}"
            )
        page_footer_found = False
        page_footer_centered = False
        for section in report.document.sections:
            for footer_paragraph in section.footer.paragraphs:
                if footer_paragraph._p.xpath(".//w:instrText[contains(., 'PAGE')]"):
                    page_footer_found = True
                    if effective_alignment(footer_paragraph) == WD_ALIGN_PARAGRAPH.CENTER:
                        page_footer_centered = True
        if not page_footer_found:
            report.errors.append("Поле номера страницы PAGE не найдено в нижнем колонтитуле тела отчёта")
        elif not page_footer_centered:
            report.errors.append("Номер страницы должен располагаться по центру нижнего колонтитула")
        if any(
            paragraph._p.xpath(".//w:instrText[contains(., 'PAGE')]")
            for paragraph in first_section.footer.paragraphs
        ) or any(
            paragraph._p.xpath(".//w:instrText[contains(., 'PAGE')]")
            for paragraph in first_section.first_page_footer.paragraphs
        ):
            report.errors.append(
                "Титульный лист, задание, реферат и перечень сокращений считаются, "
                "но номер страницы на них не печатается (tr23-26 3.4.2)"
            )
