"""Стратегии титульного листа."""

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
from report_build.document.styles import add_centered_paragraph
from report_build.spec.textutil import abbreviate_name, required

class CoverBuilder(ReportPart):
    """Титульный лист. Макет выбирается по metadata.cover_layout."""

    def add_cover(self, document: Document, spec: dict[str, Any]) -> None:
        """Dispatch the title page by cover_layout; the body never depends on it."""
        meta = spec["metadata"]
        report_type = spec["report_type"]
        layout = str(meta.get("cover_layout", "")).strip().casefold() or ("lab" if report_type == "lab" else "course")
        if layout not in COVER_LAYOUTS:
            raise SpecError(
                "metadata.cover_layout must be one of: " + ", ".join(sorted(COVER_LAYOUTS))
            )
        if layout == "vkr_2026":
            self.add_cover_vkr_2026(document, spec)
        elif layout == "vkr_tr23":
            self.add_cover_vkr_tr23(document, spec)
        elif layout == "nir":
            self.add_cover_nir(document, spec)
        else:
            self.add_cover_student(document, spec, report_type)

    def add_first_page_footer_text(self, document: Document, text: str) -> None:
        section = document.sections[0]
        section.different_first_page_header_footer = True
        paragraph = section.first_page_footer.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(0)
        set_run_font(paragraph.add_run(text), "Times New Roman", 14)

    def add_cover_vkr_2026(self, document: Document, spec: dict[str, Any]) -> None:
        """Титульный лист ВКР по официальной форме 2026 года; институт и высшая школа — ИКНК и ВШК."""
        meta = spec["metadata"]
        ministry = str(meta.get("ministry", "Министерство науки и высшего образования Российской Федерации")).strip()
        organization_type = str(meta.get("organization_type", "Федеральное государственное автономное образовательное учреждение высшего образования")).strip()
        institution = str(meta.get("institution", "Санкт-Петербургский политехнический университет Петра Великого")).strip()
        institute = str(meta.get("institute", "Институт компьютерных наук и кибербезопасности")).strip()
        department = str(meta.get("department", "Высшая школа кибербезопасности")).strip()
        add_centered_paragraph(document, ministry, bold=True)
        add_centered_paragraph(document, f"{organization_type} «{institution.strip('«»')}»", bold=True)
        add_centered_paragraph(document, institute, bold=True)
        department_paragraph = add_centered_paragraph(document, department, bold=True)
        department_paragraph.paragraph_format.space_after = Pt(64)
        add_centered_paragraph(document, "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА", bold=True, after=8)
        work_title = required(meta, "title").strip("«»\"")
        add_centered_paragraph(document, f"на тему: «{work_title}»", bold=True, after=8)
        spacer = add_centered_paragraph(document, "")
        spacer.paragraph_format.space_before = Pt(40)
        spacer.paragraph_format.space_after = Pt(40)
        student_course = str(meta.get("student_course", "")).strip() or "____"
        student_group = str(meta.get("student_group", "")).strip() or "____"
        student_name = str(meta.get("student_name", "")).strip()
        advisor_name = str(meta.get("advisor_name", meta.get("reviewer_name", ""))).strip()
        advisor_position = str(meta.get("advisor_position", meta.get("reviewer_position", ""))).strip()
        student_display = abbreviate_name(student_name, spaced_initials=False) if student_name else "________________________"
        advisor_display = abbreviate_name(advisor_name, spaced_initials=True) if advisor_name else "________________________"
        block_rows = [
            f"Выполнил(а): студент(ка) {student_course} курса, группы {student_group}",
            f"Ф.И.О.: {student_display}",
            "Подпись: ________________________",
            "",
            "Научный руководитель:",
            advisor_position or "ученая степень, должность: ________________________",
            f"Ф.И.О.: {advisor_display}",
            "Оценка: _________________________",
            "Подпись: ________________________",
            "",
            "Нормоконтроль:",
            "Подпись: ________________________",
        ]
        for row_text in block_rows:
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(2)
            if row_text:
                set_run_font(paragraph.add_run(row_text), "Times New Roman", 14)
        city = required(meta, "city")
        year = required(meta, "year")
        self.add_first_page_footer_text(document, f"{city} {year}")

    def add_cover_vkr_tr23(self, document: Document, spec: dict[str, Any]) -> None:
        """Титульный лист ВКР по Приложению А пособия tr23-26; институт и высшая школа — ИКНК и ВШК."""
        meta = spec["metadata"]
        ministry = str(meta.get("ministry", "Министерство науки и высшего образования Российской Федерации")).strip()
        institution = str(meta.get("institution", "Санкт-Петербургский политехнический университет Петра Великого")).strip()
        institute = str(meta.get("institute", "Институт компьютерных наук и кибербезопасности")).strip()
        department = str(meta.get("department", "Высшая школа кибербезопасности")).strip()
        add_centered_paragraph(document, ministry)
        add_centered_paragraph(document, institution.upper())
        add_centered_paragraph(document, institute)
        department_paragraph = add_centered_paragraph(document, department, bold=True)
        department_paragraph.paragraph_format.space_after = Pt(36)
        approver_name = str(meta.get("approver_name", "")).strip()
        approver_display = abbreviate_name(approver_name, spaced_initials=True) if approver_name else ""
        admission_rows = [
            str(meta.get("approver_role", "Работа допущена к защите")),
            str(meta.get("approver_position", "Директор ВШК")),
            f"_____________ {approver_display}".rstrip(),
            "«___» ________________ 20__ г.",
        ]
        for row_index, row_text in enumerate(admission_rows):
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(12 if row_index == len(admission_rows) - 1 else 2)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(self.profile.content_width_cm()), WD_TAB_ALIGNMENT.RIGHT)
            set_run_font(paragraph.add_run(f"\t{row_text}"), "Times New Roman", 14)
        add_centered_paragraph(document, "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА", bold=True, after=2)
        add_centered_paragraph(document, "вид ВКР (работа бакалавра, дипломный проект, дипломная работа, магистерская диссертация)", size=12, after=16)
        work_title = required(meta, "title").strip("«»\"")
        add_centered_paragraph(document, work_title.upper(), bold=True, after=8)
        direction = str(meta.get("direction", "")).strip()
        if direction:
            add_centered_paragraph(document, f"по направлению подготовки {direction}", after=4)
        profile_name = str(meta.get("profile_name", "")).strip()
        if profile_name:
            add_centered_paragraph(document, f"направленность (профиль) {profile_name}", after=8)
        spacer = add_centered_paragraph(document, "")
        spacer.paragraph_format.space_before = Pt(20)
        student_group = str(meta.get("student_group", "")).strip() or "____"
        student_name = str(meta.get("student_name", "")).strip()
        advisor_name = str(meta.get("advisor_name", meta.get("reviewer_name", ""))).strip()
        advisor_position = str(meta.get("advisor_position", meta.get("reviewer_position", ""))).strip()
        student_display = abbreviate_name(student_name, spaced_initials=True) if student_name else "И.О. Фамилия"
        advisor_display = abbreviate_name(advisor_name, spaced_initials=True) if advisor_name else "И.О. Фамилия"
        normocontrol_name = str(meta.get("normocontrol_name", "")).strip()
        normocontrol_display = abbreviate_name(normocontrol_name, spaced_initials=True) if normocontrol_name else "И.О. Фамилия"
        signer_rows = [
            ("Выполнил", ""),
            (f"студент гр. {student_group}", student_display),
            ("Руководитель", ""),
            (advisor_position or "должность, ученая степень", advisor_display),
            ("Консультант по нормоконтролю", normocontrol_display),
        ]
        for left, right in signer_rows:
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(2)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(self.profile.content_width_cm() * 0.55), WD_TAB_ALIGNMENT.CENTER)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(self.profile.content_width_cm()), WD_TAB_ALIGNMENT.RIGHT)
            if left:
                set_run_font(paragraph.add_run(left), "Times New Roman", 14)
            if right:
                set_run_font(paragraph.add_run(f"\t<подпись>\t{right}"), "Times New Roman", 12, italic=True)
        add_centered_paragraph(document, "")
        city = required(meta, "city")
        year = required(meta, "year")
        add_centered_paragraph(document, city, after=2)
        add_centered_paragraph(document, year)

    def add_cover_nir(self, document: Document, spec: dict[str, Any]) -> None:
        """Титульный лист отчёта о НИР по ГОСТ 7.32-2017 6.10."""
        meta = spec["metadata"]
        ministry = str(meta.get("ministry", "Министерство науки и высшего образования Российской Федерации")).strip()
        organization_type = str(meta.get("organization_type", "Федеральное государственное автономное образовательное учреждение высшего образования")).strip()
        institution = str(meta.get("institution", "Санкт-Петербургский политехнический университет Петра Великого")).strip()
        short_name = str(meta.get("short_name", "СПбПУ")).strip()
        add_centered_paragraph(document, ministry)
        add_centered_paragraph(document, organization_type.upper())
        add_centered_paragraph(document, institution.upper())
        add_centered_paragraph(document, f"({short_name.upper()})")
        for label, value in (
            ("УДК", str(meta.get("udc", "")).strip()),
            ("Рег. N НИОКТР", str(meta.get("reg_nioktr", "")).strip()),
            ("Рег. N ИКРБС", str(meta.get("reg_ikrbs", "")).strip()),
        ):
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(2)
            set_run_font(paragraph.add_run(f"{label} {value}" if value else label), "Times New Roman", 14)
        approver_name = str(meta.get("approver_name", "")).strip()
        approver_display = abbreviate_name(approver_name, spaced_initials=True) if approver_name else ""
        approve_rows = [
            "УТВЕРЖДАЮ",
            str(meta.get("approver_position", "")) or "________________",
            f"_____________ {approver_display}".rstrip(),
            "«__» ____________ 20__ г.",
        ]
        for row_index, row_text in enumerate(approve_rows):
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(10 if row_index == len(approve_rows) - 1 else 2)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(self.profile.content_width_cm()), WD_TAB_ALIGNMENT.RIGHT)
            set_run_font(paragraph.add_run(f"\t{row_text}"), "Times New Roman", 14)
        add_centered_paragraph(document, "ОТЧЕТ", bold=True, after=2)
        add_centered_paragraph(document, "О НАУЧНО-ИССЛЕДОВАТЕЛЬСКОЙ РАБОТЕ", bold=True, after=16)
        work_title = required(meta, "title").strip("«»\"")
        add_centered_paragraph(document, work_title.upper(), bold=True, after=4)
        stage = str(meta.get("report_stage", "")).strip()
        if stage:
            add_centered_paragraph(document, f"({stage})", after=8)
        leader_name = str(meta.get("leader_name", meta.get("advisor_name", ""))).strip()
        leader_position = str(meta.get("leader_position", "")).strip()
        for row_text in ("Руководитель НИР,", leader_position or "________________"):
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(2)
            set_run_font(paragraph.add_run(row_text), "Times New Roman", 14)
        paragraph = document.add_paragraph(style="Report Cover")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.first_line_indent = Cm(0)
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_after = Pt(2)
        paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(self.profile.content_width_cm()), WD_TAB_ALIGNMENT.RIGHT)
        set_run_font(paragraph.add_run("________________"), "Times New Roman", 14)
        if leader_name:
            set_run_font(paragraph.add_run(f"\t{abbreviate_name(leader_name, spaced_initials=True)}"), "Times New Roman", 14)
        bottom = add_centered_paragraph(document, "")
        bottom.paragraph_format.space_before = Pt(100)
        city = required(meta, "city")
        year = required(meta, "year")
        add_centered_paragraph(document, f"{city} {year}")

    def add_cover_student(self, document: Document, spec: dict[str, Any], report_type: str) -> None:
        meta = spec["metadata"]

        ministry = str(meta.get("ministry", "Министерство науки и высшего образования Российской Федерации")).strip()
        organization_type = str(
            meta.get(
                "organization_type",
                "Федеральное государственное автономное образовательное учреждение высшего образования",
            )
        ).strip()
        institution = str(meta.get("institution", "Санкт-Петербургский политехнический университет Петра Великого")).strip()
        institute = str(meta.get("institute", "Институт компьютерных наук и кибербезопасности")).strip()
        department = str(meta.get("department", "Высшая школа кибербезопасности")).strip()

        if ministry:
            add_centered_paragraph(document, ministry)
        if organization_type:
            add_centered_paragraph(document, organization_type)
        if institution:
            add_centered_paragraph(document, f"«{institution.strip('«»')}»")
        if institute:
            add_centered_paragraph(document, institute)
        if department:
            department_paragraph = add_centered_paragraph(document, department, bold=True)
            department_paragraph.paragraph_format.space_after = Pt(154)

        if report_type == "lab":
            number = str(meta.get("work_number", "")).strip()
            work_label = "Лабораторная работа" + (f" № {number}" if number else "")
        else:
            work_label = "Курсовая работа"
        add_centered_paragraph(document, work_label, size=14, bold=True, after=16)
        work_title = required(meta, "title").strip("«»\"")
        add_centered_paragraph(document, f"«{work_title}»", size=14, bold=True, after=8)
        variant = str(meta.get("variant", "")).strip()
        if variant:
            add_centered_paragraph(document, f"Вариант {variant}", size=14, after=4)
        discipline = required(meta, "discipline").strip("«»\"")
        discipline_paragraph = add_centered_paragraph(document, f"по дисциплине «{discipline}»", size=14)
        discipline_paragraph.paragraph_format.space_after = Pt(86)

        student_role = str(meta.get("student_role", "Выполнил")).strip()
        student_status = str(meta.get("student_status", "студент")).strip()
        student_display_name = str(meta.get("student_display_name", "")).strip() or abbreviate_name(
            required(meta, "student_name"), spaced_initials=False
        )
        reviewer_display_name = str(meta.get("reviewer_display_name", "")).strip() or abbreviate_name(
            required(meta, "reviewer_name"), spaced_initials=True
        )
        reviewer_position = required(meta, "reviewer_position")
        detail_rows = [
            (student_role, "", 0.0),
            (f"{student_status} гр. {required(meta, 'student_group')}", student_display_name, 0.0),
        ]
        for left, right, after in detail_rows:
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(after)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(self.profile.content_width_cm()), WD_TAB_ALIGNMENT.RIGHT)
            if left:
                run = paragraph.add_run(left)
                set_run_font(run, "Times New Roman", 14)
            if right:
                run = paragraph.add_run(f"\t{right}")
                set_run_font(run, "Times New Roman", 14)

        signature = add_centered_paragraph(document, "<подпись>", size=12, after=66)
        for run in signature.runs:
            run.font.italic = True

        reviewer_rows = [
            (str(meta.get("reviewer_role", "Проверил:")).strip(), "", 0.0),
            (reviewer_position, reviewer_display_name, 0.0),
        ]
        for left, right, after in reviewer_rows:
            paragraph = document.add_paragraph(style="Report Cover")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_after = Pt(after)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Cm(self.profile.content_width_cm()), WD_TAB_ALIGNMENT.RIGHT)
            if left:
                set_run_font(paragraph.add_run(left), "Times New Roman", 14)
            if right:
                set_run_font(paragraph.add_run(f"\t{right}"), "Times New Roman", 14)

        signature = add_centered_paragraph(document, "<подпись>", size=12, after=93)
        for run in signature.runs:
            run.font.italic = True

        add_centered_paragraph(document, required(meta, "city"), size=14)
        add_centered_paragraph(document, required(meta, "year"), size=14)
