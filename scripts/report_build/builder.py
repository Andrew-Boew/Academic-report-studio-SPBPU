"""Конвейер сборки DOCX.

Порядок один и тот же для лабораторной, курсовой, ВКР и НИР.
Меняются только стратегии: CoverBuilder по макету титула и
ContentBuilder по типу блока. Тело документа от макета не зависит.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Any, Iterable

from docx import Document

from report_build.constants import FRONT_MATTER_TYPES
from report_build.covers.builder import CoverBuilder
from report_build.content.equations import EquationBuilder
from report_build.errors import SpecError
from report_build.content.figures import FigureBuilder
from report_build.content.front import FrontMatterBuilder
from report_build.content.listings import ListingBuilder
from report_build.document.markup import MarkupWriter
from report_build.profile import FormatProfile
from report_build.content.prose import ProseBuilder
from report_build.spec.validator import SpecValidator
from report_build.state import BuildState
from report_build.document.styles import StyleBuilder
from report_build.content.tables import TableBuilder


class ContentBuilder:
    """Стратегия блока: поле type выбирает метод, который пишет этот блок."""

    def __init__(self, ctx: ReportBuilder) -> None:
        self.ctx = ctx

    @property
    def state(self) -> BuildState:
        return self.ctx.state

    def add(self, document: Document, blocks: Iterable[dict[str, Any]]) -> None:
        handlers = {
            "heading": self.ctx.prose.add_heading,
            "paragraph": self.ctx.prose.add_paragraph_block,
            "note": self.ctx.prose.add_note,
            "abbreviations": self.ctx.front.add_abbreviations,
            "abstract": self.ctx.front.add_abstract,
            "assignment": self.ctx.front.add_assignment,
            "bullet_list": lambda doc, block: self.ctx.prose.add_list(doc, block, False),
            "numbered_list": lambda doc, block: self.ctx.prose.add_list(doc, block, True),
            "figure": self.ctx.figures.add_figure,
            "figure_placeholder": self.ctx.figures.add_figure_placeholder,
            "table": self.ctx.tables.add_table,
            "equation": self.ctx.equations.add_equation,
            "code": self.ctx.listings.add_code,
            "appendix_heading": self.ctx.front.add_appendix_heading,
            "bibliography": self.ctx.front.add_bibliography,
        }
        for index, block in enumerate(blocks):
            if not isinstance(block, dict):
                raise SpecError(f"content[{index}] must be an object")
            block_type = block.get("type")
            if block_type == "page_break":
                document.add_page_break()
                self.state.clear_gap()
                continue
            handler = handlers.get(block_type)
            if handler is None:
                raise SpecError(f"Unsupported content block type at index {index}: {block_type}")
            handler(document, block)
            if block_type == "heading":
                self.state.last_block_kind = "heading"
            elif block_type in {"paragraph", "bullet_list", "numbered_list", "note"}:
                self.state.last_block_kind = "text"
            else:
                self.state.last_block_kind = "object"


class ReportBuilder:
    """Собирает один DOCX из уже проверенной спеки."""

    def __init__(self, spec: dict[str, Any]) -> None:
        self.spec = spec
        self.state = BuildState()
        self.profile = FormatProfile()
        self.document: Document | None = None
        self.markup = MarkupWriter(self)
        self.styles = StyleBuilder(self)
        self.covers = CoverBuilder(self)
        self.prose = ProseBuilder(self)
        self.front = FrontMatterBuilder(self)
        self.figures = FigureBuilder(self)
        self.tables = TableBuilder(self)
        self.equations = EquationBuilder(self)
        self.listings = ListingBuilder(self)
        self.validator = SpecValidator(self)
        self.blocks = ContentBuilder(self)

    def build(self, output: Path) -> None:
        spec = self.spec
        self.validator.validate_spec(spec)
        settings = spec.get("document") or {}
        requested_profile = str(settings.get("format_profile", "unified")).strip()
        self.profile.activate(requested_profile)
        self.state.footnotes.clear()
        self.state.pending_gap = False
        self.state.last_block_kind = None
        document = Document()
        self.document = document
        # Continuous physical numbering from the title page; the title page number
        # stays hidden via a different first-page footer, so the first visible
        # number always matches the real rendered position without hardcoding.
        self.styles.configure_document(document, int(settings.get("page_number_start", 1) or 1))

        properties = document.core_properties
        # Keep document properties completely empty for clean metadata and privacy
        properties.title = str(settings.get("title", ""))
        properties.subject = str(settings.get("subject", ""))
        properties.author = str(settings.get("author", ""))
        properties.last_modified_by = str(settings.get("author", ""))
        properties.keywords = ""
        properties.comments = ""
        properties.category = ""
        properties.content_status = ""
        properties.identifier = ""
        now_utc = dt.datetime.now(dt.timezone.utc)
        properties.created = now_utc
        properties.modified = now_utc

        document._report_blocks = spec["content"]  # type: ignore[attr-defined]
        self.state.equation_numbers.clear()
        for block in spec["content"]:
            if isinstance(block, dict) and block.get("type") == "equation":
                number = str(block.get("number", "")).strip()
                if number:
                    self.state.equation_numbers.add(number)

        self.covers.add_cover(document, spec)
        content = [dict(block) for block in spec["content"]]
        front_matter = [
            block for block in content
            if block.get("type") in FRONT_MATTER_TYPES
            and str(block.get("position", "front")).strip().casefold() != "body"
        ]
        body_blocks = [
            block for block in content
            if not (block.get("type") in FRONT_MATTER_TYPES
                    and str(block.get("position", "front")).strip().casefold() != "body")
        ]
        for block in front_matter:
            self.blocks.add(document, [block])
        include_toc = bool(settings.get("include_toc", True))
        if include_toc:
            self.styles.start_body_section(document)
            levels = max(1, min(3, int(settings.get("toc_levels", 3))))
            self.front.add_toc(document, levels)
        if not include_toc:
            self.styles.start_body_section(document)
        if include_toc and body_blocks:
            if body_blocks[0].get("type") == "heading":
                body_blocks[0]["page_break_before"] = True
            else:
                body_blocks.insert(0, {"type": "page_break"})
        self.blocks.add(document, body_blocks)
        self.markup.attach_footnotes(document)

        output.parent.mkdir(parents=True, exist_ok=True)
        document.save(output)


def build(spec: dict[str, Any], output: Path) -> None:
    ReportBuilder(spec).build(output)
