"""Общие предикаты и имена закладок проверки. Состояния документа не держат."""

from __future__ import annotations

import re
from typing import Any

from docx.oxml.ns import qn

PLACEHOLDER = re.compile(r"\{\{|\}\}|\b(?:TODO|TBD|FIXME)\b|\[вставить[^\]]*\]", re.I)
OBJECT_NUMBER = r"(?:[А-ЯЁ]\.\d+|\d+(?:\.\d+)*)"
TABLE_CAPTION = re.compile(rf"^Таблица\s+({OBJECT_NUMBER})\s+–\s+\S.*$", re.I)
TABLE_CONTINUATION = re.compile(rf"^(Продолжение|Окончание)\s+таблицы\s+({OBJECT_NUMBER})$", re.I)
FIGURE_CAPTION = re.compile(rf"^Рисунок\s+({OBJECT_NUMBER})\s+–\s+\S.*$", re.I)
LISTING_CAPTION = re.compile(rf"^Листинг\s+({OBJECT_NUMBER})\s+–\s+\S.*$", re.I)
MANUAL_HEADING_NUMBER = re.compile(r"^\s*\d+(?:\.\d+)*[.)]?\s+")
FIGURE_REFERENCE = re.compile(
    r"\bрис(?:унок|унка|унке|унках|унков|унку|\.?)\s+"
    r"((?:[А-ЯЁ]\.)?\d+)",
    re.IGNORECASE,
)
TABLE_REFERENCE = re.compile(r"\bтабл\w*\s+((?:[А-ЯЁ]\.)?\d+(?:\.\d+)*)", re.I)
LISTING_REFERENCE = re.compile(r"\bлистинг\w*\s+((?:[А-ЯЁ]\.)?\d+(?:\.\d+)*)", re.I)
EQUATION_REFERENCE = re.compile(r"\(((?:[А-ЯЁ]\.)?\d+(?:\.\d+)*)\)")
CYRILLIC_BOOKMARK_LETTERS = {
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Е": "E",
    "Ж": "ZH", "З": "Z", "И": "I", "К": "K", "Л": "L", "М": "M",
    "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T",
    "У": "U", "Ф": "F", "Х": "KH", "Ц": "TS", "Ч": "CH", "Ш": "SH",
    "Щ": "SCH", "Э": "E", "Ю": "YU", "Я": "YA",
}


def near(value: float, expected: float, tolerance: float = 0.03) -> bool:
    return abs(value - expected) <= tolerance


def style_font_name(style) -> str | None:
    return style.font.name


def xml_text(element) -> str:
    return "".join(node.text or "" for node in element.findall(".//" + qn("w:t"))).strip()


def is_layout_table(table_element) -> bool:
    """Identify the borderless cover-details table, which is not a numbered data table."""
    borders = table_element.find("./" + qn("w:tblPr") + "/" + qn("w:tblBorders"))
    if borders is None or len(borders) == 0:
        return False
    values = [border.get(qn("w:val"), "") for border in borders]
    return bool(values) and all(value in {"nil", "none"} for value in values)


def is_code_fragment(table_element) -> bool:
    descriptor = table_element.find("./" + qn("w:tblPr") + "/" + qn("w:tblCaption"))
    return descriptor is not None and descriptor.get(qn("w:val")) == "code-fragment"


def is_figure_placeholder(table_element) -> bool:
    descriptor = table_element.find("./" + qn("w:tblPr") + "/" + qn("w:tblCaption"))
    return descriptor is not None and descriptor.get(qn("w:val")) == "figure-placeholder"


def effective_alignment(paragraph):
    return paragraph.alignment if paragraph.alignment is not None else paragraph.style.paragraph_format.alignment


def warn_numeric_gaps(numbers: list[str], label: str, warnings: list[str]) -> None:
    if numbers and all(number.isdigit() for number in numbers):
        values = [int(number) for number in numbers]
        expected = list(range(1, len(values) + 1))
        if values != expected:
            warnings.append(f"Нарушена последовательная сквозная нумерация {label}: {values}")


def object_bookmark_name(kind: str, number: str) -> str:
    normalized = number.strip().upper()
    match = re.fullmatch(r"([А-ЯЁ])\.(\d+)", normalized)
    if match:
        letter = CYRILLIC_BOOKMARK_LETTERS.get(match.group(1), f"CYR{ord(match.group(1))}")
        return f"_Report{kind}_{letter}_{match.group(2)}"
    return f"_Report{kind}_" + re.sub(r"[^A-Z0-9_]", "_", normalized)


def figure_bookmark_name(number: str) -> str:
    normalized = number.strip().upper()
    match = re.fullmatch(r"([А-ЯЁ])\.(\d+)", normalized)
    if match:
        letter = CYRILLIC_BOOKMARK_LETTERS.get(match.group(1), f"CYR{ord(match.group(1))}")
        return f"_ReportFigure_{letter}_{match.group(2)}"
    return "_ReportFigure_" + re.sub(r"[^A-Z0-9_]", "_", normalized)
