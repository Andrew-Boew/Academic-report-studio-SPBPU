"""План удаления черновиков после проверенного отчёта.

Сначала план отказывается от опасных путей, потом показывает их,
и только с --apply удаляет.
"""

from __future__ import annotations

import shutil
from pathlib import Path


def inside(child: Path, parent: Path) -> bool:
    try:
        child.relative_to(parent)
        return True
    except ValueError:
        return False


class CleanupPlan:
    """Список того, что можно удалить, не задев итоговый DOCX."""

    def __init__(self, final: Path, work: Path, drafts: list[Path]) -> None:
        self.final = final
        self.work = work
        self.drafts = drafts
        self._guard()

    def _guard(self) -> None:
        if not self.final.is_file() or self.final.suffix.lower() != ".docx" or self.final.stat().st_size == 0:
            raise SystemExit("Final DOCX is missing or empty; cleanup refused")
        if self.work == self.work.parent or len(self.work.parts) < 4:
            raise SystemExit(f"Work directory is too broad: {self.work}")
        if inside(self.final, self.work):
            raise SystemExit("Final DOCX must be outside the disposable work directory")
        for draft in self.drafts:
            if draft == self.final:
                raise SystemExit("A draft path resolves to the final DOCX")

    def targets(self) -> list[Path]:
        return [path for path in [self.work, *self.drafts] if path.exists()]

    def preview(self) -> list[Path]:
        found = self.targets()
        for target in found:
            print(f"WOULD_DELETE\t{target}")
        return found

    def apply(self) -> list[Path]:
        found = self.targets()
        for target in found:
            print(f"DELETE\t{target}")
        for target in found:
            if target.is_dir() and not target.is_symlink():
                shutil.rmtree(target)
            else:
                target.unlink()
        return found
