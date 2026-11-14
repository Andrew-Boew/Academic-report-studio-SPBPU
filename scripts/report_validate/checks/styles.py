"""Стили абзаца, заголовков и многоуровневая нумерация."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class StyleCheck:
    """Normal, Heading 1–6, структурные стили и номера."""

    def apply(self, report) -> None:
        normal = report.document.styles["Normal"]
        font_name = style_font_name(normal)
        if font_name != "Times New Roman":
            report.errors.append(f"Стиль Normal: шрифт {font_name!r} вместо Times New Roman")
        if normal.font.size is None or not near(normal.font.size.pt, 14.0, 0.1):
            report.errors.append("Стиль Normal: размер должен быть 14 pt")
        if normal.paragraph_format.first_line_indent is None or not near(normal.paragraph_format.first_line_indent.cm, 1.25):
            report.errors.append("Стиль Normal: абзацный отступ должен быть 1,25 см")
        if normal.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
            report.errors.append("Стиль Normal: требуется выравнивание по ширине")
        spacing = normal.paragraph_format.line_spacing
        if not isinstance(spacing, float) or not near(spacing, 1.5, 0.01):
            report.errors.append("Стиль Normal: межстрочный интервал должен быть 1,5")

        expected_heading_sizes = {
            "Heading 1": 18.0, "Heading 2": 16.0, "Heading 3": 14.0,
            "Heading 4": 14.0, "Heading 5": 14.0, "Heading 6": 14.0,
        }
        expected_heading_spacing = {
            "Heading 1": (0.0, 18.0),
            "Heading 2": (12.0, 6.0),
            "Heading 3": (12.0, 6.0),
            "Heading 4": (12.0, 6.0),
            "Heading 5": (12.0, 6.0),
            "Heading 6": (12.0, 6.0),
        }
        for name in expected_heading_sizes:
            style = report.document.styles[name]
            expected_size = expected_heading_sizes[name]
            if style.font.name != "Times New Roman" or style.font.size is None or not near(style.font.size.pt, expected_size, 0.1):
                report.errors.append(f"Стиль {name}: требуется Times New Roman {expected_size:g} pt")
            if style.paragraph_format.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                report.errors.append(f"Стиль {name}: нумерованные заголовки должны быть выровнены слева")
            first_indent = style.paragraph_format.first_line_indent
            first_indent_cm = first_indent.cm if first_indent is not None else 0.0
            left_indent = style.paragraph_format.left_indent
            left_indent_cm = left_indent.cm if left_indent is not None else 0.0
            if not near(first_indent_cm, 0.0) or not near(left_indent_cm, 0.0):
                report.errors.append(
                    f"Стиль {name}: прямые отступы должны быть 0 см; общий отступ 1,25 см задаёт нумерация Word"
                )
            before_pt, after_pt = expected_heading_spacing[name]
            if style.paragraph_format.space_before is None or not near(style.paragraph_format.space_before.pt, before_pt, 0.1):
                report.errors.append(f"Стиль {name}: интервал перед должен быть {before_pt:g} pt")
            if style.paragraph_format.space_after is None or not near(style.paragraph_format.space_after.pt, after_pt, 0.1):
                report.errors.append(f"Стиль {name}: интервал после должен быть {after_pt:g} pt")
            line_spacing = style.paragraph_format.line_spacing
            if not isinstance(line_spacing, float) or not near(line_spacing, 1.0, 0.01):
                report.errors.append(f"Стиль {name}: межстрочный интервал должен быть одинарным")
        if report.document.styles["Heading 3"].font.italic:
            report.errors.append("Стиль Heading 3 не должен быть курсивом")
        if not report.document.styles["Heading 6"].font.italic:
            report.errors.append("Стиль Heading 6 должен быть курсивом")

        heading_num_ids: list[str] = []
        for level, name in enumerate(("Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6")):
            ppr = report.document.styles[name]._element.get_or_add_pPr()
            num_pr = ppr.find(qn("w:numPr"))
            ilvl = num_pr.find(qn("w:ilvl")) if num_pr is not None else None
            num_id = num_pr.find(qn("w:numId")) if num_pr is not None else None
            if ilvl is None or ilvl.get(qn("w:val")) != str(level) or num_id is None:
                report.errors.append(f"Стиль {name} не связан с уровнем {level + 1} многоуровневой нумерации Word")
                continue
            heading_num_ids.append(num_id.get(qn("w:val"), ""))
        if heading_num_ids and len(set(heading_num_ids)) != 1:
            report.errors.append("Heading 1–6 должны использовать одну многоуровневую нумерацию Word")
        if heading_num_ids and len(set(heading_num_ids)) == 1:
            numbering = report.document.part.numbering_part.element
            num_node = next(
                (node for node in numbering.findall(qn("w:num")) if node.get(qn("w:numId")) == heading_num_ids[0]),
                None,
            )
            abstract_ref = num_node.find(qn("w:abstractNumId")) if num_node is not None else None
            abstract_id = abstract_ref.get(qn("w:val")) if abstract_ref is not None else None
            abstract = next(
                (
                    node for node in numbering.findall(qn("w:abstractNum"))
                    if node.get(qn("w:abstractNumId")) == abstract_id
                ),
                None,
            )
            for level in range(6):
                lvl = next(
                    (
                        node for node in abstract.findall(qn("w:lvl"))
                        if node.get(qn("w:ilvl")) == str(level)
                    ),
                    None,
                ) if abstract is not None else None
                indent = lvl.find(qn("w:pPr") + "/" + qn("w:ind")) if lvl is not None else None
                label_node = lvl.find(qn("w:lvlText")) if lvl is not None else None
                expected_labels = (
                    "%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4", "%1.%2.%3.%4.%5", "%1.%2.%3.%4.%5.%6",
                )
                if label_node is None or label_node.get(qn("w:val")) != expected_labels[level]:
                    report.errors.append(
                        f"Уровень {level + 1}: неверный формат номера для профиля {report.format_profile}"
                    )
                if (
                    indent is None
                    or indent.get(qn("w:left")) != "1417"
                    or indent.get(qn("w:hanging")) != "708"
                ):
                    report.errors.append(
                        f"Уровень {level + 1} нумерации заголовков должен начинать номер на 1,25 см и текст на 2,50 см"
                    )

        if "Report Structural Heading" not in [style.name for style in report.document.styles]:
            report.errors.append("Не найден стиль ненумерованных структурных заголовков")
        else:
            structural_style = report.document.styles["Report Structural Heading"]
            structural_num_pr = structural_style._element.get_or_add_pPr().find(qn("w:numPr"))
            if structural_num_pr is not None:
                report.errors.append("Структурные заголовки не должны наследовать нумерацию глав")

        for name in ("Report Cover", "Report Caption", "Report Table Caption", "Report Equation"):
            if name not in [style.name for style in report.document.styles]:
                report.errors.append(f"Не найден стиль {name}")
                continue
            style = report.document.styles[name]
            if style.font.name != "Times New Roman" or style.font.size is None or not near(style.font.size.pt, 14.0, 0.1):
                if name != "Report Equation":
                    report.errors.append(f"Стиль {name}: требуется Times New Roman 14 pt")
            if name in {"Report Caption", "Report Table Caption", "Report Equation"}:
                line_spacing = style.paragraph_format.line_spacing
                before = style.paragraph_format.space_before
                after = style.paragraph_format.space_after
                expected_spacing = report.expected_caption_spacing if name in {"Report Caption", "Report Table Caption"} else 1.5
                if not isinstance(line_spacing, float) or not near(line_spacing, expected_spacing, 0.01):
                    report.errors.append(f"Стиль {name}: межстрочный интервал должен быть {expected_spacing:g}")
                expected_before = 12.0 if name == "Report Table Caption" else 0.0
                expected_after = 18.0 if name == "Report Caption" else 3.0 if name == "Report Table Caption" else 0.0
                if before is None or not near(before.pt, expected_before, 0.1):
                    report.errors.append(f"Стиль {name}: интервал перед должен быть {expected_before:g} pt")
                if after is None or not near(after.pt, expected_after, 0.1):
                    report.errors.append(f"Стиль {name}: интервал после должен быть {expected_after:g} pt")
