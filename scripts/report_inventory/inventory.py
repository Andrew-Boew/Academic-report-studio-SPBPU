"""Инвентарь каталога: стратегия выбирается по типу файла."""

from __future__ import annotations

import mimetypes
from pathlib import Path
from typing import Any

from report_inventory.support import (
    TEXT_EXTENSIONS,
    collect,
    inspect_csv,
    inspect_docx,
    inspect_image,
    inspect_json,
    inspect_pdf,
    inspect_text,
    inspect_xlsx,
    sha256,
)

class FileProbe:
    """Один вид файла. Совпавший пробник дополняет карточку, остальные не вызываются."""

    def matches(self, path: Path) -> bool:
        raise NotImplementedError

    def describe(self, path: Path) -> dict[str, Any]:
        raise NotImplementedError

class CsvProbe(FileProbe):
    def matches(self, path: Path) -> bool:
        return path.suffix.lower() in {".csv", ".tsv"}

    def describe(self, path: Path) -> dict[str, Any]:
        return inspect_csv(path)

class WorkbookProbe(FileProbe):
    def matches(self, path: Path) -> bool:
        return path.suffix.lower() in {".xlsx", ".xlsm"}

    def describe(self, path: Path) -> dict[str, Any]:
        return inspect_xlsx(path)

class JsonProbe(FileProbe):
    def matches(self, path: Path) -> bool:
        return path.suffix.lower() == ".json"

    def describe(self, path: Path) -> dict[str, Any]:
        return inspect_json(path)

class DocxProbe(FileProbe):
    def matches(self, path: Path) -> bool:
        return path.suffix.lower() == ".docx"

    def describe(self, path: Path) -> dict[str, Any]:
        return inspect_docx(path)

class PdfProbe(FileProbe):
    def matches(self, path: Path) -> bool:
        return path.suffix.lower() == ".pdf"

    def describe(self, path: Path) -> dict[str, Any]:
        return inspect_pdf(path)

class ImageProbe(FileProbe):
    def matches(self, path: Path) -> bool:
        return path.suffix.lower() in {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp",
        }

    def describe(self, path: Path) -> dict[str, Any]:
        return inspect_image(path)

class TextProbe(FileProbe):
    def matches(self, path: Path) -> bool:
        return path.suffix.lower() in TEXT_EXTENSIONS

    def describe(self, path: Path) -> dict[str, Any]:
        return inspect_text(path)

class Inventory:
    """Обходит каталог и для каждого файла берёт первый подходящий пробник."""

    probes: tuple[FileProbe, ...] = (
        CsvProbe(),
        WorkbookProbe(),
        JsonProbe(),
        DocxProbe(),
        PdfProbe(),
        ImageProbe(),
        TextProbe(),
    )

    def describe_file(self, path: Path, root: Path) -> dict[str, Any]:
        suffix = path.suffix.lower()
        item: dict[str, Any] = {
            "path": str(path.resolve()),
            "relative_path": str(path.relative_to(root)),
            "extension": suffix,
            "mime_type": mimetypes.guess_type(path.name)[0],
            "size_bytes": path.stat().st_size,
            "sha256": sha256(path),
        }
        for probe in self.probes:
            if probe.matches(path):
                item.update(probe.describe(path))
                break
        else:
            item["kind"] = "other"
        if path.name == ".report-profile.json":
            item["kind"] = "report_profile"
        elif path.name == ".report-subjects.json":
            item["kind"] = "subject_profiles"
        return item

    def build(self, root: Path) -> dict[str, Any]:
        items = []
        errors = []
        for path in collect(root):
            try:
                items.append(self.describe_file(path, root))
            except Exception as exc:
                errors.append({"path": str(path), "error": f"{type(exc).__name__}: {exc}"})
        return {
            "root": str(root),
            "file_count": len(items),
            "total_size_bytes": sum(item["size_bytes"] for item in items),
            "files": items,
            "errors": errors,
            "privacy_note": (
                "Headers and structural metadata only; dataset rows are not copied into this inventory."
            ),
        }
