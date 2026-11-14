"""Содержание."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class TocCheck:
    """Отступы, табуляция и номера страниц содержания."""

    def apply(self, report) -> None:
        toc_expected_indents = report.expected_toc_indents
        for style_name, expected_indent in toc_expected_indents.items():
            if style_name not in [style.name for style in report.document.styles]:
                continue
            indent = report.document.styles[style_name].paragraph_format.left_indent
            actual_indent = indent.cm if indent is not None else 0.0
            if not near(actual_indent, expected_indent):
                report.errors.append(
                    f"Стиль {style_name}: левый отступ {actual_indent:.2f} см вместо {expected_indent:.2f} см"
                )
            line_spacing = report.document.styles[style_name].paragraph_format.line_spacing
            if line_spacing is None or not near(float(line_spacing), 1.5, 0.05):
                report.errors.append(
                    f"Стиль {style_name}: межстрочный интервал содержания должен быть полуторным (1,5)"
                )

        toc_paragraphs = [
            paragraph for paragraph in report.document.paragraphs
            if paragraph.style and paragraph.style.name in toc_expected_indents
        ]
        for paragraph in toc_paragraphs:
            style_name = paragraph.style.name
            expected_indent = toc_expected_indents[style_name]
            direct_indent = paragraph.paragraph_format.left_indent
            style_indent = report.document.styles[style_name].paragraph_format.left_indent
            effective_indent = (
                direct_indent.cm
                if direct_indent is not None
                else (style_indent.cm if style_indent is not None else 0.0)
            )
            if not near(effective_indent, expected_indent):
                report.errors.append(
                    f"Строка {style_name}: левый отступ {effective_indent:.2f} см вместо {expected_indent:.2f} см"
                )
            direct_first = paragraph.paragraph_format.first_line_indent
            effective_first = direct_first.cm if direct_first is not None else 0.0
            if not near(effective_first, 0.0):
                report.errors.append(f"Строка {style_name}: отступ первой строки содержания должен быть 0 см")
        update_fields = report.document.settings._element.find(qn("w:updateFields"))
        update_value = update_fields.get(qn("w:val"), "true").lower() if update_fields is not None else ""
        if toc_paragraphs and update_value not in {"false", "0", "off"}:
            report.errors.append(
                "После материализации содержания автоматическое updateFields должно быть отключено, чтобы Word не показывал нули"
            )
        if report.has_toc and not toc_paragraphs and update_value not in {"true", "1", "on"}:
            report.errors.append("Черновое содержание должно запрашивать обновление полей перед материализацией")
        for paragraph in toc_paragraphs:
            tabs = paragraph._p.xpath("./w:pPr/w:tabs/w:tab[@w:val='right'][@w:leader='dot']")
            if not tabs:
                report.errors.append(f"Строка содержания не имеет правой табуляции с точечным заполнителем: {paragraph.text!r}")
                continue
            expected_toc_tab_pos = "9354"
            if tabs[0].get(qn("w:pos")) != expected_toc_tab_pos:
                expected_width_label = "16,5 см"
                report.errors.append(
                    f"Правый табулятор содержания должен находиться на рабочей ширине {expected_width_label} "
                    f"(позиция {expected_toc_tab_pos} twips)"
                )
            text_nodes = paragraph._p.xpath(".//w:t/text()")
            if not text_nodes or not text_nodes[-1].strip().isdigit():
                report.errors.append(f"В строке содержания отсутствует номер страницы: {paragraph.text!r}")
            elif int(text_nodes[-1].strip()) <= 0:
                report.errors.append(f"В строке содержания указан нулевой номер страницы: {paragraph.text!r}")
