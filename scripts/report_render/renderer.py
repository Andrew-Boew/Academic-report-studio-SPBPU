"""Конвейер рендера: DOCX в PDF, затем страницы в PNG.

Преобразователь LibreOffice один. Растеризатор — стратегия:
сначала PyMuPDF, и только если его нет, pdftoppm.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path


class LibreOffice:
    """Ищет soffice и конвертирует DOCX в PDF во временном профиле."""

    def find(self) -> Path:
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

    def to_pdf(self, docx: Path, work: Path, verbose: bool) -> Path:
        soffice = self.find()
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


class Rasterizer:
    """Один способ превратить PDF в page-N.png."""

    def available(self) -> bool:
        raise NotImplementedError

    def render(self, pdf: Path, output_dir: Path, dpi: int) -> int:
        raise NotImplementedError


class FitzRasterizer(Rasterizer):
    def available(self) -> bool:
        try:
            import fitz
        except ImportError:
            return False
        self._fitz = fitz
        return True

    def render(self, pdf: Path, output_dir: Path, dpi: int) -> int:
        fitz = self._fitz
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


class PdftoppmRasterizer(Rasterizer):
    def available(self) -> bool:
        self._binary = os.environ.get("PDFTOPPM_BIN") or shutil.which("pdftoppm")
        return bool(self._binary)

    def render(self, pdf: Path, output_dir: Path, dpi: int) -> int:
        prefix = output_dir / "page"
        result = subprocess.run(
            [self._binary, "-png", "-r", str(dpi), str(pdf), str(prefix)],
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


class DocxRenderer:
    """DOCX → PDF → PNG. Первый доступный растеризатор и есть стратегия."""

    def __init__(self) -> None:
        self.office = LibreOffice()
        self.rasterizers: tuple[Rasterizer, ...] = (FitzRasterizer(), PdftoppmRasterizer())

    def rasterize(self, pdf: Path, output_dir: Path, dpi: int) -> int:
        for rasterizer in self.rasterizers:
            if rasterizer.available():
                return rasterizer.render(pdf, output_dir, dpi)
        raise RuntimeError(
            "Для PNG-рендера нужен PyMuPDF (requirements-render.txt) или pdftoppm из Poppler."
        )

    def render(
        self,
        source: Path,
        output: Path,
        dpi: int,
        *,
        emit_pdf: bool,
        verbose: bool,
        work: Path,
    ) -> int:
        pdf = self.office.to_pdf(source, work, verbose)
        pages = self.rasterize(pdf, output, dpi)
        if emit_pdf:
            shutil.copy2(pdf, output / f"{source.stem}.pdf")
        return pages
