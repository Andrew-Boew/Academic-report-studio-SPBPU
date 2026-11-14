#!/usr/bin/env python3
"""Render DOCX to page PNG files on Windows, macOS, or Linux.

LibreOffice performs the DOCX-to-PDF conversion. PyMuPDF rasterizes the PDF.
The script uses an isolated LibreOffice profile and does not modify the input.
"""

from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

from report_render.renderer import DocxRenderer


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--output_dir", required=True, type=Path)
    parser.add_argument("--dpi", type=int, default=150)
    parser.add_argument("--emit_pdf", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    source = args.input.expanduser().resolve()
    if source.suffix.lower() != ".docx" or not source.is_file():
        parser.error(f"DOCX не найден: {source}")
    if args.dpi < 72 or args.dpi > 600:
        parser.error("--dpi должен находиться в диапазоне 72..600")
    output = args.output_dir.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    for old_page in output.glob("page-*.png"):
        old_page.unlink()

    try:
        with tempfile.TemporaryDirectory(prefix="academic-report-render-") as temp_name:
            pages = DocxRenderer().render(
                source,
                output,
                args.dpi,
                emit_pdf=args.emit_pdf,
                verbose=args.verbose,
                work=Path(temp_name),
            )
    except Exception as exc:
        print(f"render_docx: {exc}", file=sys.stderr)
        return 1
    print(f"Отрендеровано страниц: {pages}; каталог: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
