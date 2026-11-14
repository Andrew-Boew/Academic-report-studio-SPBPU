"""Чистые помощники текста и реквизитов. Состояния сборки не держат."""

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

def clean_text(value: Any, *, field: str) -> str:
    text = str(value if value is not None else "").strip()
    if PLACEHOLDER_PATTERN.search(text):
        raise SpecError(f"Placeholder detected in {field}: {text[:80]}")
    return text

def required(metadata: dict[str, Any], key: str) -> str:
    value = str(metadata.get(key, "")).strip()
    if not value:
        raise SpecError(f"Missing required metadata field: {key}")
    return value

def abbreviate_name(value: str, *, spaced_initials: bool) -> str:
    """Convert a full Russian name to the surname-and-initials title-page form."""
    value = value.strip()
    if "." in value:
        return value
    parts = [part for part in value.split() if part]
    if len(parts) < 2:
        return value
    separator = " " if spaced_initials else ""
    initials = separator.join(f"{part[0].upper()}." for part in parts[1:])
    return f"{parts[0]} {initials}"

def capitalized_label(value: str) -> str:
    value = value.strip()
    return value[:1].upper() + value[1:] if value else value

def count_spec_objects(blocks: list[dict[str, Any]]) -> tuple[int, int, int, int]:
    figures = tables = sources = appendices = 0
    for block in blocks:
        if not isinstance(block, dict):
            continue
        block_type = block.get("type")
        if block_type in {"figure", "figure_placeholder"}:
            figures += 1
        elif block_type == "table":
            tables += 1
        elif block_type == "bibliography":
            sources += len(block.get("items", []) or [])
        elif block_type == "appendix_heading":
            appendices += 1
    return figures, tables, sources, appendices

def keep_paragraph_lines(paragraph) -> None:
    paragraph.paragraph_format.widow_control = True

def keep_previous_reference(document: Document, *, max_characters: int = 180) -> None:
    """Keep a short reference paragraph with the following caption/object."""
    for paragraph in reversed(document.paragraphs):
        if not paragraph.text.strip():
            continue
        if paragraph.style and paragraph.style.name in {"Report Caption", "Report Table Caption", "Report Figure Annotation"}:
            return
        if paragraph.text.strip().startswith(("Рисунок ", "Таблица ", "Листинг ")):
            return
        # Keeping a long paragraph with a following figure/table can move the
        # whole paragraph to the next page and leave half of the prior page
        # empty. Only a genuinely short lead-in is kept with the object.
        if len(paragraph.text) <= max_characters:
            paragraph.paragraph_format.keep_with_next = True
        return
