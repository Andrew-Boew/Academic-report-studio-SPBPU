"""Каталог пунктов содержания по JSON-спеке."""

from __future__ import annotations

import re

from report_toc.support import structural_heading

class EntryCatalog:
    """Собирает заголовки, источники и приложения в порядке спеки."""

    def collect(self, spec: dict) -> list[dict]:
        levels = max(1, min(3, int((spec.get("document") or {}).get("toc_levels", 3))))
        # The unified body profile (GOST 7.32-2017 / tr23-26) displays heading
        # numbers without a trailing period.
        heading_suffix = ""
        entries: list[dict] = []
        counters = [0, 0, 0]
        for block in spec.get("content", []):
            kind = block.get("type")
            if kind == "heading":
                level = int(block.get("level", 1))
                text = str(block.get("text", "")).strip()
                if text and level <= levels:
                    if structural_heading(block):
                        display_text = text
                    else:
                        counters[level - 1] += 1
                        for index in range(level, 3):
                            counters[index] = 0
                        if any(value == 0 for value in counters[:level]):
                            raise ValueError(f"Invalid heading hierarchy before: {text}")
                        display_text = (
                            f"{'.'.join(str(value) for value in counters[:level])}{heading_suffix} {text}"
                        )
                    entries.append({"text": text, "display_text": display_text, "level": level})
            elif kind == "bibliography":
                text = str(block.get("title", "Список использованных источников")).strip()
                entries.append({"text": text, "display_text": text, "level": 1})
            elif kind == "appendix_heading":
                text = f"Приложение {str(block.get('letter', '')).strip().upper()}"
                entries.append({"text": text, "display_text": text, "level": 1})
        return entries

