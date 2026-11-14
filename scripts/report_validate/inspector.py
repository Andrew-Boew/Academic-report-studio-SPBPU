"""Проверка DOCX как цепочка независимых проверок.

ReportInspector.inspect открывает файл и вызывает проверки в фиксированном
порядке. Каждая проверка только добавляет ошибки в общий отчёт.
Порядок совпадает с прежней функцией inspect, поэтому список сообщений тот же.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

from docx import Document

from report_validate.checks.cover import CoverCheck
from report_validate.checks.links import LinkCheck
from report_validate.checks.objects import CaptionCheck, ReferenceCheck
from report_validate.checks.package import PackageCheck
from report_validate.checks.page import PageCheck
from report_validate.checks.profile import ProfileCheck
from report_validate.checks.structure import (
    FrontMatterCheck,
    HeadingCheck,
    ProseCheck,
    SourceCheck,
)
from report_validate.checks.styles import StyleCheck
from report_validate.checks.tables import BodyObjectCheck
from report_validate.checks.toc import TocCheck


class ReportInspector:
    """Один проход по документу. Проверки не вызывают друг друга."""

    checks = (
        ProfileCheck,
        PageCheck,
        StyleCheck,
        CoverCheck,
        PackageCheck,
        TocCheck,
        LinkCheck,
        HeadingCheck,
        FrontMatterCheck,
        SourceCheck,
        ProseCheck,
        CaptionCheck,
        ReferenceCheck,
        BodyObjectCheck,
    )

    def inspect(self, path: Path) -> dict[str, Any]:
        self.path = path
        self.errors: list[str] = []
        self.warnings: list[str] = []
        self.document = None
        if not self._open():
            return {
                "path": str(self.path),
                "errors": self.errors,
                "warnings": self.warnings,
                "passed": False,
            }
        for check in self.checks:
            check().apply(self)
        return self._result()

    def _open(self) -> bool:
        if self.path.suffix.lower() != ".docx":
            self.errors.append("Итоговый файл должен иметь расширение .docx")
            return False
        if not self.path.is_file():
            self.errors.append("Файл не найден")
            return False
        if not zipfile.is_zipfile(self.path):
            self.errors.append("Файл не является корректным контейнером DOCX/ZIP")
            return False
        try:
            self.document = Document(self.path)
        except Exception as exc:
            self.errors.append(f"DOCX не открывается: {type(exc).__name__}: {exc}")
            return False
        return True

    def _result(self) -> dict[str, Any]:
        return {
            "path": str(self.path),
            "report_type": self.report_type,
            "format_profile": self.format_profile,
            "size_bytes": self.path.stat().st_size,
            "sections": len(self.document.sections),
            "paragraphs": sum(bool(p.text.strip()) for p in self.document.paragraphs),
            "tables": len(self.document.tables),
            "content_tables": self.content_tables,
            "figure_numbers": self.figure_numbers,
            "table_numbers": self.table_numbers,
            "listing_numbers": self.listing_numbers,
            "equation_numbers": self.equation_numbers,
            "inline_images": len(self.document.inline_shapes),
            "bibliography_sources": len(self.bibliography_paragraphs),
            "external_source_links": self.seen_external_links,
            "errors": self.errors,
            "warnings": self.warnings,
            "passed": not self.errors,
        }


def inspect(path: Path) -> dict[str, Any]:
    return ReportInspector().inspect(path)
