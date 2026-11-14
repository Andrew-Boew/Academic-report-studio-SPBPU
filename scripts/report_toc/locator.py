"""Поиск страницы каждого пункта по уже отрисованному PDF."""

from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader

from report_toc.support import page_lines, squashed

class PageLocator:
    """Ставит номер страницы по точному совпадению строки, без нечёткого поиска."""

    def locate(self, entries: list[dict], pdf_path: Path, page_start: int) -> list[dict]:
        reader = PdfReader(str(pdf_path))
        lines_by_page = [page_lines(page) for page in reader.pages]
        for entry in entries:
            visible = entry.get("display_text", entry["text"])
            target = squashed(visible)
            matches = []
            for physical_index, line_variants in enumerate(lines_by_page):
                # Some pypdf releases prepend the right-footer page number to the
                # first extracted line (for example, ``2`` + ``1 ЦЕЛЬ РАБОТЫ``).
                # Accept that single, precisely predictable variant without using
                # fuzzy matching that could confuse a TOC entry with the heading.
                displayed_page = str(physical_index + page_start)
                first_line_with_footer = squashed(f"{displayed_page}{visible}")
                for lines in line_variants:
                    exact_line = any(squashed(line) == target for line in lines)
                    wrapped_heading = any(
                        squashed("".join(lines[start:start + count])) == target
                        for count in (2, 3)
                        for start in range(0, max(0, len(lines) - count + 1))
                    )
                    if exact_line or wrapped_heading or (
                        lines and squashed(lines[0]) == first_line_with_footer
                    ):
                        matches.append(physical_index)
                        break
            if not matches:
                raise ValueError(f"Heading not found as a standalone rendered line: {visible}")
            physical_index = matches[-1]
            entry["page"] = physical_index + page_start
        return entries

