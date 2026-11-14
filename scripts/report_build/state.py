"""Состояние одной сборки: сноски, зазор после объекта, номера формул."""

from __future__ import annotations

from dataclasses import dataclass, field

from docx.shared import Pt

from report_build.constants import OBJECT_GAP_PT

@dataclass
class BuildState:
    pending_gap: bool = False
    pending_gap_pt: float = OBJECT_GAP_PT
    last_block_kind: str | None = None
    equation_numbers: set[str] = field(default_factory=set)
    footnotes: list[str] = field(default_factory=list)

    def request_gap(self, points: float | None = None) -> None:
        """Ask for a gap above the next body paragraph."""
        self.pending_gap = True
        self.pending_gap_pt = OBJECT_GAP_PT if points is None else points

    def apply_gap(self, paragraph_format) -> None:
        if self.pending_gap:
            self.pending_gap = False
            paragraph_format.space_before = Pt(self.pending_gap_pt)

    def clear_gap(self) -> None:
        self.pending_gap = False
        self.pending_gap_pt = OBJECT_GAP_PT

    def consume_before(self, minimum: float) -> float:
        """Keep an object's own space_before, unless a previous object asked for more."""
        value = minimum
        if self.pending_gap:
            value = max(value, self.pending_gap_pt)
        self.clear_gap()
        return value
