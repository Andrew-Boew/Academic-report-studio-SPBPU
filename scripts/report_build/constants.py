"""Профиль оформления и образцы, общие для всех шагов сборки."""

from __future__ import annotations

import re
from typing import Any

FORMAT_PROFILES: dict[str, dict[str, Any]] = {
    "unified": {
        "left_margin_cm": 3.0,
        "right_margin_cm": 1.5,
        "top_margin_cm": 2.0,
        "bottom_margin_cm": 1.5,
        "heading_labels": (
            "%1",
            "%1.%2",
            "%1.%2.%3",
            "%1.%2.%3.%4",
            "%1.%2.%3.%4.%5",
            "%1.%2.%3.%4.%5.%6",
        ),
        "uppercase_structural": True,
        "table_caption_alignment": "left",
        "table_font_size": 12.0,
        "table_line_spacing": 1.0,
        "caption_line_spacing": 1.0,
        "toc_indents_cm": {1: 0.0, 2: 0.5, 3: 1.0},
        "manual_table_continuations": True,
        "automatic_hyphenation": True,
        "strict_units": True,
        "strict_appendices": True,
        "dash_bullet_lists": True,
    },
}
# ГОСТ 7.32-2017 and the SPbPU 2023 NIR/VKR manual define ONE body format for
# every work type (лабораторная, курсовая, ВКР, отчёт по НИР). The work type
# changes only the title page and the set of structural elements, never the
# body rules. Historical profile names are accepted and mapped to `unified`.
LEGACY_FORMAT_PROFILES = {"spbpu": "unified", "student_default": "unified"}
ACTIVE_FORMAT_PROFILE_NAME = "unified"
ACTIVE_FORMAT_PROFILE = FORMAT_PROFILES["unified"]
MAX_CONTENT_WIDTH_CM = 16.5
# Empty paragraphs are never used as vertical spacers.
BLANK_LINE_PT = 0.0
# Equations keep a moderate 8 pt gap. Figures and tables follow the NIR
# template: 12 pt above a figure, 6 pt under the image, 18 pt after the
# figure caption; 12 pt before a table caption, 3 pt after it, and 18 pt
# before the paragraph that follows the table. List items use 3 pt.
OBJECT_GAP_PT = 8.0
FIGURE_BEFORE_PT = 12.0
FIGURE_AFTER_PT = 6.0
FIGURE_CAPTION_AFTER_PT = 18.0
TABLE_CAPTION_BEFORE_PT = 12.0
TABLE_CAPTION_AFTER_PT = 3.0
TABLE_FOLLOWING_PT = 18.0
OMML_NAMESPACES = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
PLACEHOLDER_PATTERN = re.compile(r"\{\{|\}\}|\b(?:TODO|TBD|FIXME)\b|\[вставить[^\]]*\]", re.I)
MANUAL_HEADING_NUMBER = re.compile(r"^\s*\d+(?:\.\d+)*[.)]?\s+")
CITATION_PATTERN = re.compile(r"\[(\d+)\]")
FIGURE_REFERENCE_PATTERN = re.compile(
    r"(?P<label>\bрис(?:унок|унка|унке|унках|унков|унку|\.?)\s+)"
    r"(?P<number>(?:[А-ЯЁ]\.)?\d+)",
    re.IGNORECASE,
)
FOOTNOTE_PATTERN = re.compile(r"\[\^([^\]\[]+)\]")
# Prose references to numbered objects; every one becomes an internal hyperlink.
TABLE_REFERENCE_PATTERN = re.compile(r"\bтабл\w*\s+((?:[А-ЯЁ]\.)?\d+(?:\.\d+)*)", re.I)
LISTING_REFERENCE_PATTERN = re.compile(r"\bлистинг\w*\s+((?:[А-ЯЁ]\.)?\d+(?:\.\d+)*)", re.I)
EQUATION_REFERENCE_PATTERN = re.compile(r"\(((?:[А-ЯЁ]\.)?\d+(?:\.\d+)*)\)")
# Inline math in prose: $...$ carries a LaTeX expression that becomes a real
# inline OMML object (Cambria Math) instead of styled plain text.
INLINE_MATH_PATTERN = re.compile(r"\$([^$]+)\$")
CYRILLIC_BOOKMARK_LETTERS = {
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Е": "E",
    "Ж": "ZH", "З": "Z", "И": "I", "К": "K", "Л": "L", "М": "M",
    "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T",
    "У": "U", "Ф": "F", "Х": "KH", "Ц": "TS", "Ч": "CH", "Ш": "SH",
    "Щ": "SCH", "Э": "E", "Ю": "YU", "Я": "YA",
}
SOURCE_KINDS = {
    "article",
    "book",
    "conference_paper",
    "dataset",
    "official_documentation",
    "official_web_resource",
    "standard",
}
FRONT_MATTER_TYPES = {"abstract", "assignment", "abbreviations"}
COVER_LAYOUTS = {"lab", "course", "vkr_2026", "vkr_tr23", "nir"}

UNIT_PATTERN = re.compile(
    r"(?<=\d)[ \t]+(?=(?:%|мм|см|дм|км|м|мг|г|кг|мл|л|мс|с|мин|ч|"
    r"Гц|кГц|МГц|ГГц|В|кВ|А|мА|Вт|кВт|Па|кПа|МПа|байт|Кбайт|Мбайт|Гбайт)\b)",
    re.IGNORECASE,
)
# Signs raised above the line (°, ') follow the number without a space
# (tr23-26 3.10.1): 120°, 15'.
DEGREE_SPACE_PATTERN = re.compile(r"(?<=\d)[ \u00a0\t]+(?=°)")

