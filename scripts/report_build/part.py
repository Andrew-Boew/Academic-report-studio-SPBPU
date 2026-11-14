"""Общий доступ шага сборки к профилю и состоянию.

Это не иерархия отчётов: классы не подменяют друг друга,
а читают один ReportBuilder, который их создал.
"""

from __future__ import annotations

class ReportPart:
    def __init__(self, ctx: ReportBuilder) -> None:
        self.ctx = ctx

    @property
    def state(self) -> BuildState:
        return self.ctx.state

    @property
    def profile(self) -> FormatProfile:
        return self.ctx.profile
