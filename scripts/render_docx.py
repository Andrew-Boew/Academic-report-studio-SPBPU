#!/usr/bin/env python3
"""Render DOCX to page PNG files on Windows, macOS, or Linux.

LibreOffice performs the DOCX-to-PDF conversion. PyMuPDF rasterizes the PDF.
The script uses an isolated LibreOffice profile and does not modify the input.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_soffice() -> Path:
    configured = os.environ.get("LIBREOFFICE_BIN", "").strip()
    candidates: list[Path] = []
    if configured:
        candidates.append(Path(configured).expanduser())
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            candidates.append(Path(found))
    program_files = [
        os.environ.get("PROGRAMFILES"),
        os.environ.get("PROGRAMFILES(X86)"),
    ]
    candidates.extend(
        [
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/bin/soffice"),
            Path("/snap/bin/libreoffice"),
        ]
    )
    for root in program_files:
        if root:
            candidates.append(Path(root) / "LibreOffice" / "program" / "soffice.exe")
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    raise RuntimeError(
        "LibreOffice не найден. Установите LibreOffice или задайте путь в "
        "переменной LIBREOFFICE_BIN; без него визуальная проверка DOCX невозможна."
    )


def convert_to_pdf(docx: Path, work: Path, verbose: bool) -> Path:
    soffice = find_soffice()
    profile = work / "lo-profile"
    profile.mkdir(parents=True, exist_ok=True)
    command = [
        str(soffice),
        "--headless",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(work),
        str(docx),
    ]
    if verbose:
        print("[render_docx] " + " ".join(command))
    result = subprocess.run(command, capture_output=True, text=True, timeout=180)
    expected = work / f"{docx.stem}.pdf"
    if result.returncode != 0 or not expected.is_file() or expected.stat().st_size == 0:
        details = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
        raise RuntimeError(f"LibreOffice не создал PDF (код {result.returncode}).\n{details}")
    return expected


def rasterize(pdf: Path, output_dir: Path, dpi: int) -> int:
    try:
        import fitz
    except ImportError:
        fitz = None
    if fitz is not None:
        scale = dpi / 72.0
        matrix = fitz.Matrix(scale, scale)
        document = fitz.open(pdf)
        try:
            for index, page in enumerate(document, 1):
                pixmap = page.get_pixmap(matrix=matrix, alpha=False)
                pixmap.save(output_dir / f"page-{index}.png")
            return len(document)
        finally:
            document.close()

    pdftoppm = os.environ.get("PDFTOPPM_BIN") or shutil.which("pdftoppm")
    if pdftoppm:
        prefix = output_dir / "page"
        result = subprocess.run(
            [pdftoppm, "-png", "-r", str(dpi), str(pdf), str(prefix)],
            capture_output=True,
            text=True,
            timeout=180,
        )
        pages = sorted(output_dir.glob("page-*.png"))
        if result.returncode == 0 and pages:
            for index, page in enumerate(pages, start=1):
                normalized = output_dir / f"page-{index}.png"
                if page != normalized:
                    page.replace(normalized)
            return len(pages)
        details = "\n".join(part for part in (result.stdout.strip(), result.stderr.strip()) if part)
        raise RuntimeError(f"pdftoppm не создал PNG (код {result.returncode}).\n{details}")

    raise RuntimeError(
        "Для PNG-рендера нужен PyMuPDF (requirements-render.txt) или pdftoppm из Poppler."
    )


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
            pdf = convert_to_pdf(source, Path(temp_name), args.verbose)
            pages = rasterize(pdf, output, args.dpi)
            if args.emit_pdf:
                shutil.copy2(pdf, output / f"{source.stem}.pdf")
    except Exception as exc:
        print(f"render_docx: {exc}", file=sys.stderr)
        return 1
    print(f"Отрендеровано страниц: {pages}; каталог: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
