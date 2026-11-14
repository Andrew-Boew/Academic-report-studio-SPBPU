"""Активный профиль оформления одной сборки.

Раньше профиль лежал в глобальных переменных модуля и менялся
на время build(). Теперь он живёт в объекте сборки.
"""

from __future__ import annotations

from typing import Any

from report_build.constants import (
    ACTIVE_FORMAT_PROFILE,
    ACTIVE_FORMAT_PROFILE_NAME,
    FORMAT_PROFILES,
    LEGACY_FORMAT_PROFILES,
)

class FormatProfile:
    def __init__(self) -> None:
        self.name = ACTIVE_FORMAT_PROFILE_NAME
        self.data = ACTIVE_FORMAT_PROFILE

    def get(self, name: str) -> Any:
        return self.data[name]

    def content_width_cm(self) -> float:
        return 21.0 - float(self.get("left_margin_cm")) - float(self.get("right_margin_cm"))

    def activate(self, requested: str) -> None:
        self.name = LEGACY_FORMAT_PROFILES.get(requested, requested)
        self.data = FORMAT_PROFILES[self.name]
