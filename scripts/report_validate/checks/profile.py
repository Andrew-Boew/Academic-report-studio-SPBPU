"""Ожидания единого профиля оформления."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class ProfileCheck:
    """Читает профиль из свойств документа и фиксирует допуски."""

    def apply(self, report) -> None:
        keywords = report.document.core_properties.keywords or ""
        marker = "academic-report-format-profile:"
        report.format_profile = keywords.split(marker, 1)[1].split()[0].strip() if marker in keywords else "unified"
        report.format_profile = {"spbpu": "unified", "student_default": "unified"}.get(report.format_profile, report.format_profile)
        if report.format_profile != "unified":
            report.errors.append(f"Неизвестный профиль оформления в свойствах DOCX: {report.format_profile}")
            report.format_profile = "unified"
        # The unified body profile (GOST 7.32-2017 / SPbPU tr23-26) applies to every
        # work type; the legacy spbpu values below are therefore always expected.
        report.strict_spbpu = True
        report.expected_right_margin = 1.5
        report.expected_page_start = "1"
        report.expected_caption_spacing = 1.0
        report.expected_toc_indents = {"TOC 1": 0.0, "TOC 2": 0.5, "TOC 3": 1.0}
        report.expected_table_size_pt = 12.0
