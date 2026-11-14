"""Сборка академического DOCX.

Читать архитектуру с ReportBuilder.build: там весь порядок документа.
CoverBuilder и ContentBuilder только выбирают, кто выполняет конкретный шаг.
"""

from report_build.builder import ReportBuilder, build
from report_build.errors import SpecError

__all__ = ["ReportBuilder", "SpecError", "build"]
