"""Титульный лист."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class CoverCheck:
    """Вид работы и подписи на титуле."""

    def apply(self, report) -> None:
        first_page_text = "\n".join(p.text for p in report.document.paragraphs[:35])
        category = (report.document.core_properties.category or "").strip()
        category_match = re.match(r"^(lab|course)\s+(lab|course|vkr_2026|vkr_tr23|nir)$", category)
        if category_match:
            report.report_type = category_match.group(1)
            report.cover_layout = category_match.group(2)
        else:
            if re.search(r"ВЫПУСКНА\s+КВАЛИФИКАЦИОННА\s+РАБОТА", first_page_text, re.I):
                report.cover_layout = "vkr_2026"
            elif re.search(r"О\s+НАУЧНО-ИССЛЕДОВАТЕЛЬСКО\s+РАБОТЕ", first_page_text, re.I):
                report.cover_layout = "nir"
            else:
                report.cover_layout = "course" if re.search(r"КУРСОВАЯ\s+РАБОТА", first_page_text, re.I) else "lab"
            report.report_type = "lab" if report.cover_layout == "lab" else "course"
        if "политехнический университет" not in first_page_text.lower():
            report.errors.append("Не найден обязательный титульный блок университета")
        expected_label = {
            "lab": r"ЛАБОРАТОРНАЯ\s+РАБОТА",
            "course": r"КУРСОВАЯ\s+РАБОТА",
            "vkr_2026": r"ВЫПУСКНАЯ\s+КВАЛИФИКАЦИОННАЯ\s+РАБОТА",
            "vkr_tr23": r"ВЫПУСКНАЯ\s+КВАЛИФИКАЦИОННАЯ\s+РАБОТА",
            "nir": r"О\s+НАУЧНО-ИССЛЕДОВАТЕЛЬСКОЙ\s+РАБОТЕ",
        }[report.cover_layout]
        if not re.search(expected_label, first_page_text, re.I):
            report.errors.append("На титульном листе не указан тип работы по выбранному титульному макету")
        work_label_paragraphs = [
            p for p in report.document.paragraphs
            if re.search(expected_label, p.text, re.I)
        ]
        if work_label_paragraphs:
            runs = [run for run in work_label_paragraphs[0].runs if run.text.strip()]
            if not runs or any(run.font.size is None or not near(run.font.size.pt, 14.0, 0.1) for run in runs):
                report.errors.append("Тип работы на титульном листе должен иметь размер 14 pt")
            if any(run.font.bold is not True for run in runs):
                report.errors.append("Тип работы на титульном листе должен быть полужирным")
        if report.cover_layout in {"lab", "course"}:
            signature_paragraphs = [p for p in report.document.paragraphs[:35] if p.text.strip() == "<подпись>"]
            if len(signature_paragraphs) != 2:
                report.errors.append("На титульном листе должны быть две строки <подпись>: для студента и проверяющего")
            for paragraph in signature_paragraphs:
                if effective_alignment(paragraph) != WD_ALIGN_PARAGRAPH.CENTER:
                    report.errors.append("Строка <подпись> должна быть выровнена по центру")
                runs = [run for run in paragraph.runs if run.text.strip()]
                if any(run.font.size is None or not near(run.font.size.pt, 12.0, 0.1) or run.font.italic is not True for run in runs):
                    report.errors.append("Строка <подпись> должна быть Times New Roman 12 pt курсивом")
