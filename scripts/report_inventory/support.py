"""Чтение файлов для инвентаря. В отчёт не попадают строки данных."""

from __future__ import annotations

import csv
import hashlib
import json
import os
from pathlib import Path
from typing import Any

IGNORED_DIRS = {
    ".git", ".idea", ".vscode", "__pycache__", "node_modules", "output", "renders", "tmp",
}
TEXT_EXTENSIONS = {
    ".c", ".cc", ".cpp", ".css", ".h", ".hpp", ".html", ".ipynb",
    ".java", ".js", ".jsx", ".md", ".m", ".py", ".r", ".rs", ".sh",
    ".sql", ".swift", ".tex", ".ts", ".tsx", ".txt", ".yaml", ".yml",
}
MAX_HASH_BYTES = 250 * 1024 * 1024
PROFILE_FILENAMES = {".report-profile.json", ".report-subjects.json"}

def sha256(path: Path) -> str | None:
    if path.stat().st_size > MAX_HASH_BYTES:
        return None
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

def read_text(path: Path, limit: int = 256_000) -> tuple[str, str]:
    raw = path.read_bytes()[:limit]
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8-replace"

def inspect_csv(path: Path) -> dict[str, Any]:
    text, encoding = read_text(path)
    lines = text.splitlines()
    sample = "\n".join(lines[:30])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ","
    rows = list(csv.reader(lines[:6], delimiter=delimiter))
    header = rows[0] if rows else []
    return {
        "kind": "table",
        "encoding": encoding,
        "delimiter": delimiter,
        "estimated_rows": max(0, len(lines) - 1),
        "columns": len(header),
        "headers": header,
    }

def inspect_xlsx(path: Path) -> dict[str, Any]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        return {"kind": "workbook", "warning": f"openpyxl unavailable: {exc}"}
    workbook = load_workbook(path, read_only=True, data_only=False)
    sheets = []
    for sheet in workbook.worksheets:
        header = []
        if sheet.max_row:
            header = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
        sheets.append(
            {
                "name": sheet.title,
                "rows": sheet.max_row,
                "columns": sheet.max_column,
                "headers": header,
                "state": sheet.sheet_state,
            }
        )
    workbook.close()
    return {"kind": "workbook", "sheets": sheets}

def inspect_json(path: Path) -> dict[str, Any]:
    text, encoding = read_text(path, limit=5_000_000)
    try:
        value = json.loads(text)
    except json.JSONDecodeError as exc:
        return {"kind": "json", "encoding": encoding, "error": str(exc)}
    result: dict[str, Any] = {"kind": "json", "encoding": encoding}
    if isinstance(value, list):
        result["root_type"] = "array"
        result["items"] = len(value)
        if value and isinstance(value[0], dict):
            result["keys"] = sorted(map(str, value[0].keys()))
    elif isinstance(value, dict):
        result["root_type"] = "object"
        result["keys"] = sorted(map(str, value.keys()))
    else:
        result["root_type"] = type(value).__name__
    return result

def inspect_docx(path: Path) -> dict[str, Any]:
    try:
        from docx import Document
    except ImportError as exc:
        return {"kind": "docx", "warning": f"python-docx unavailable: {exc}"}
    document = Document(path)
    headings = [
        p.text.strip()
        for p in document.paragraphs
        if p.text.strip() and p.style and p.style.name.startswith("Heading")
    ]
    return {
        "kind": "docx",
        "sections": len(document.sections),
        "paragraphs": sum(bool(p.text.strip()) for p in document.paragraphs),
        "tables": len(document.tables),
        "inline_images": len(document.inline_shapes),
        "headings": headings[:100],
    }

def inspect_pdf(path: Path) -> dict[str, Any]:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        return {"kind": "pdf", "warning": f"pypdf unavailable: {exc}"}
    reader = PdfReader(path)
    encrypted = bool(reader.is_encrypted)
    return {"kind": "pdf", "pages": len(reader.pages), "encrypted": encrypted}

def inspect_image(path: Path) -> dict[str, Any]:
    try:
        from PIL import Image
        with Image.open(path) as image:
            return {
                "kind": "image",
                "width_px": image.width,
                "height_px": image.height,
                "mode": image.mode,
                "format": image.format,
            }
    except Exception as exc:  # Pillow reports format-specific errors.
        return {"kind": "image", "warning": str(exc)}

def inspect_text(path: Path) -> dict[str, Any]:
    text, encoding = read_text(path)
    return {
        "kind": "text_or_code",
        "encoding": encoding,
        "lines": text.count("\n") + (1 if text else 0),
        "nonempty_lines": sum(bool(line.strip()) for line in text.splitlines()),
    }

def collect(root: Path) -> list[Path]:
    paths = []
    for current_root, dirnames, filenames in os.walk(root):
        dirnames[:] = [name for name in dirnames if name not in IGNORED_DIRS and not name.startswith(".")]
        for name in filenames:
            if (name.startswith(".") and name not in PROFILE_FILENAMES) or name.startswith("~$"):
                continue
            paths.append(Path(current_root) / name)
    return sorted(paths)
