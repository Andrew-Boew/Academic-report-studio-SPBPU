"""Слияние профиля и преподавателя."""

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

def load_json_object(path: Path, label: str) -> dict[str, Any]:
    path = path.expanduser().resolve()
    if not path.is_file():
        raise SpecError(f"{label} does not exist: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SpecError(f"{label} is not valid JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise SpecError(f"{label} must be a JSON object")
    return value

def normalize_discipline(value: Any) -> str:
    return str(value or "").strip().strip("«»\"").casefold()

class MetadataMerger:
    """Общий профиль и преподаватель по дисциплине поверх metadata спеки.

    Шаг выполняется до сборки документа и не зависит от её состояния.
    """

    def merge_metadata_sources(self, 
        spec: dict[str, Any],
        profile_path: Path | None,
        subjects_path: Path | None,
    ) -> dict[str, Any]:
        current = spec.get("metadata")
        if not isinstance(current, dict):
            raise SpecError("metadata must be an object")
        merged: dict[str, Any] = {}

        if profile_path is not None:
            profile = load_json_object(profile_path, "Metadata profile")
            forbidden = sorted(key for key in profile if key.startswith("reviewer_"))
            if forbidden:
                raise SpecError(
                    "Global metadata profile must not contain discipline-specific reviewer fields: "
                    + ", ".join(forbidden)
                )
            merged.update(profile)

        discipline = current.get("discipline") or merged.get("discipline")
        if subjects_path is not None:
            subjects = load_json_object(subjects_path, "Subject profile")
            normalized = normalize_discipline(discipline)
            if not normalized:
                raise SpecError("discipline is required to select a subject-specific reviewer")
            matches = [value for key, value in subjects.items() if normalize_discipline(key) == normalized]
            if not matches:
                raise SpecError(f"No reviewer profile found for discipline: {discipline}")
            subject = matches[0]
            if not isinstance(subject, dict):
                raise SpecError(f"Reviewer profile for discipline {discipline!r} must be an object")
            allowed = {"reviewer_role", "reviewer_position", "reviewer_name"}
            unexpected = sorted(set(subject) - allowed)
            if unexpected:
                raise SpecError(
                    f"Reviewer profile for discipline {discipline!r} has unsupported fields: "
                    + ", ".join(unexpected)
                )
            merged.update(subject)

        for key, value in current.items():
            if value not in (None, ""):
                merged[key] = value
        result = dict(spec)
        result["metadata"] = merged
        return result
