"""Входная проверка спеки."""

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
from report_build.spec.textutil import clean_text, required

class SpecValidator(ReportPart):
    """Проверка JSON-спеки до того, как документ начнёт собираться."""

    def validate_spec(self, spec: dict[str, Any]) -> None:
        if spec.get("report_type") not in {"lab", "course"}:
            raise SpecError("report_type must be 'lab' or 'course'")
        if not isinstance(spec.get("metadata"), dict):
            raise SpecError("metadata must be an object")
        if not isinstance(spec.get("content"), list) or not spec["content"]:
            raise SpecError("content must be a non-empty array")
        document_settings = spec.get("document") or {}
        if not isinstance(document_settings, dict):
            raise SpecError("document must be an object")
        format_profile = str(document_settings.get("format_profile", "unified")).strip()
        format_profile = LEGACY_FORMAT_PROFILES.get(format_profile, format_profile)
        if format_profile not in FORMAT_PROFILES:
            raise SpecError(
                "document.format_profile must be one of: " + ", ".join(sorted(FORMAT_PROFILES))
            )
        code_change = document_settings.get("code_change")
        if code_change is not None:
            if not isinstance(code_change, dict):
                raise SpecError("document.code_change must be an object")
            mode = str(code_change.get("mode", "")).strip().lower()
            if mode not in {"new", "rewrite", "refactor", "modify", "none"}:
                raise SpecError("document.code_change.mode must be new, rewrite, refactor, modify, or none")
            source_keys = ("original_sources", "final_sources", "test_sources", "supporting_sources")
            declared_sources: list[Path] = []
            groups: dict[str, list[Path]] = {}
            for source_key in source_keys:
                raw_sources = code_change.get(source_key, [])
                if not isinstance(raw_sources, list):
                    raise SpecError(f"document.code_change.{source_key} must be an array")
                resolved_sources: list[Path] = []
                for raw_source in raw_sources:
                    source_path = Path(str(raw_source)).expanduser().resolve()
                    if not source_path.is_file():
                        raise SpecError(f"Declared code source does not exist: {source_path}")
                    resolved_sources.append(source_path)
                    declared_sources.append(source_path)
                groups[source_key] = resolved_sources
            if mode in {"new", "rewrite", "modify"} and not groups["final_sources"]:
                raise SpecError(f"document.code_change mode {mode!r} requires final_sources")
            if mode == "refactor" and (not groups["original_sources"] or not groups["final_sources"]):
                raise SpecError("document.code_change mode 'refactor' requires original_sources and final_sources")
            if mode == "none" and declared_sources:
                raise SpecError("document.code_change mode 'none' cannot declare source files")
            if len(set(declared_sources)) != len(declared_sources):
                raise SpecError("document.code_change lists contain a duplicate source path")
            appendix_sources: list[Path] = []
            appendix_labels: list[str] = []
            for block in spec["content"]:
                if not isinstance(block, dict) or block.get("type") != "code":
                    continue
                if str(block.get("presentation", "fragment")).strip().lower() != "appendix":
                    continue
                source_value = str(block.get("source", "")).strip()
                if source_value:
                    appendix_sources.append(Path(source_value).expanduser().resolve())
                appendix_labels.append(str(block.get("label", "")).strip())
            if mode != "none" and not any(
                isinstance(block, dict) and block.get("type") == "appendix_heading"
                for block in spec["content"]
            ):
                raise SpecError("programming report requires at least one appendix_heading")
            for source_path in declared_sources:
                if appendix_sources.count(source_path) != 1:
                    raise SpecError(
                        f"Declared code source must appear exactly once as appendix code: {source_path}"
                    )
            if mode != "none" and any(not label for label in appendix_labels):
                raise SpecError("every appendix code block requires a visible relative-path label")
        report_type = spec["report_type"]
        bibliography_blocks = [
            block for block in spec["content"]
            if isinstance(block, dict) and block.get("type") == "bibliography"
        ]
        if report_type == "course" and len(bibliography_blocks) != 1:
            raise SpecError("course content must contain exactly one bibliography with verified real sources")
        if report_type == "lab" and bibliography_blocks:
            raise SpecError("laboratory reports must not contain a bibliography or external source hyperlinks")
        bibliography_count = len(bibliography_blocks[0].get("items", [])) if bibliography_blocks else 0
        citation_numbers: list[int] = []
        for block in spec["content"]:
            if not isinstance(block, dict):
                continue
            values: list[str] = []
            if block.get("type") == "paragraph":
                values.append(str(block.get("text", "")))
            elif block.get("type") in {"bullet_list", "numbered_list"}:
                values.extend(str(item) for item in block.get("items", []))
            for value in values:
                citation_numbers.extend(int(number) for number in CITATION_PATTERN.findall(value))
        if report_type == "lab" and citation_numbers:
            raise SpecError("laboratory reports must not contain numeric source citations such as [1]")
        if report_type == "course" and any(number < 1 or number > bibliography_count for number in citation_numbers):
            raise SpecError("course-paper citation [N] refers to a missing bibliography item")
        previous_numbered_level = 0
        for index, block in enumerate(spec["content"]):
            if not isinstance(block, dict) or block.get("type") != "heading":
                continue
            level = int(block.get("level", 1))
            text = clean_text(block.get("text"), field=f"content[{index}] heading")
            if MANUAL_HEADING_NUMBER.match(text):
                raise SpecError(
                    f"content[{index}] heading contains a manual number; supply only the title text"
                )
            explicit_structural = block.get("structural") if "structural" in block else None
            inferred_structural = level == 1 and bool(re.match(
                r"^(?:СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ВЫВОДЫ|СПИСОК|ПРИЛОЖЕНИЕ)", text, re.I
            ))
            structural = bool(explicit_structural) if explicit_structural is not None else inferred_structural
            if structural:
                continue
            if level > previous_numbered_level + 1:
                raise SpecError(f"Heading hierarchy skips a level at content[{index}]: {text!r}")
            previous_numbered_level = level
        cover_layout = str((spec["metadata"] or {}).get("cover_layout", "")).strip().casefold() or (
            "lab" if spec["report_type"] == "lab" else "course"
        )
        if cover_layout not in COVER_LAYOUTS:
            raise SpecError(
                "metadata.cover_layout must be one of: " + ", ".join(sorted(COVER_LAYOUTS))
            )
        required_keys = {"title", "city", "year"}
        if cover_layout in {"lab", "course"}:
            required_keys |= {
                "discipline", "student_group", "student_name",
                "reviewer_position", "reviewer_name",
            }
        elif cover_layout in {"vkr_2026", "vkr_tr23"}:
            required_keys |= {"student_name", "student_group"}
        for key in sorted(required_keys):
            required(spec["metadata"], key)
