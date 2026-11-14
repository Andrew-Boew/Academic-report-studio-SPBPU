"""Конвейер содержания: каталог, страницы PDF, запись в DOCX."""

from __future__ import annotations

import json
from pathlib import Path

from pypdf import PdfReader

from report_toc.catalog import EntryCatalog
from report_toc.locator import PageLocator
from report_toc.patch import TocPatcher

class TocFinalizer:
    """Один проход. Каждый шаг получает результат предыдущего и не знает про CLI."""

    def run(
        self,
        source: Path,
        spec: dict,
        pdf_path: Path,
        output: Path,
        entries_json: Path | None = None,
    ) -> dict:
        document_settings = spec.get("document") or {}
        page_start = int(document_settings.get("page_number_start", 1) or 1)
        entries = EntryCatalog().collect(spec)
        entries = PageLocator().locate(entries, pdf_path, page_start)
        if not entries:
            raise SystemExit("No TOC entries found in the specification")
        page_count = len(PdfReader(str(pdf_path)).pages)
        TocPatcher().patch(source, output, entries, page_count=page_count)
        if entries_json:
            entries_json.parent.mkdir(parents=True, exist_ok=True)
            entries_json.write_text(
                json.dumps({"entries": entries, "page_count": page_count}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        return {"output": str(output), "entries": entries, "page_count": page_count}
