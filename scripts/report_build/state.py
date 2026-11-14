"""Состояние одной сборки: сноски, зазор после объекта, номера формул."""

from __future__ import annotations

from dataclasses import dataclass, field

from docx.shared import Pt

from report_build.constants import OBJECT_GAP_PT

@dataclass
class BuildState:
    pending_gap: bool = False
    last_block_kind: str | None = None
    equation_numbers: set[str] = field(default_factory=set)
    footnotes: list[str] = field(default_factory=list)

    def request_gap(self) -> None:
        """Ask for one blank-line gap above the next body paragraph."""
        self.pending_gap = True

    def apply_gap(self, paragraph_format) -> None:
        if self.pending_gap:
            self.pending_gap = False
            paragraph_format.space_before = Pt(OBJECT_GAP_PT)

    def clear_gap(self) -> None:
        self.pending_gap = False
