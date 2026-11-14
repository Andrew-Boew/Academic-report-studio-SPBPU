"""Запись найденных номеров в поле TOC документа Word."""

from __future__ import annotations

import re
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

from report_toc.support import (
    NS,
    attach_heading_bookmarks,
    entry_paragraph,
    fill_abstract_volume,
    qn,
)

class TocPatcher:
    """Подменяет черновое поле содержания готовыми строками и отключает автообновление."""

    def patch(self, 
        source: Path,
        output: Path,
        entries: list[dict],
        *,
        page_count: int,
    ) -> None:
        with zipfile.ZipFile(source) as archive:
            infos = archive.infolist()
            members = {info.filename: archive.read(info.filename) for info in infos}
        root = etree.fromstring(members["word/document.xml"])
        attach_heading_bookmarks(root, entries)
        filled = fill_abstract_volume(root, page_count)
        candidates = root.xpath(".//w:p[.//w:instrText[contains(., 'TOC')]]", namespaces=NS)
        if not candidates:
            raise ValueError("Word TOC field not found")
        old = candidates[0]
        instruction_nodes = old.xpath(".//w:instrText[contains(., 'TOC')]", namespaces=NS)
        instruction = instruction_nodes[0].text or ' TOC \\o "1-3" \\h \\z \\u '
        parent = old.getparent()
        index = parent.index(old)
        end_index = index
        siblings = list(parent)
        for candidate_index in range(index, len(siblings)):
            end_index = candidate_index
            if siblings[candidate_index].xpath(
                ".//w:fldChar[@w:fldCharType='end']", namespaces=NS
            ):
                break
        else:
            raise ValueError("End of Word TOC field not found")
        for candidate_index in range(end_index, index - 1, -1):
            parent.remove(siblings[candidate_index])
        # The unified body profile uses 16.5 cm of working width (right margin
        # 1.5 cm): the page-number tab sits at 9354 twips and hierarchy indents
        # are 0 / 0.5 / 1.0 cm (tr23-26 2.5.1).
        toc_tab_pos = "9354"
        toc_left_twips = {1: "0", 2: "284", 3: "567"}
        for offset, entry in enumerate(entries):
            parent.insert(index + offset, entry_paragraph(
                entry,
                first=offset == 0,
                last=offset == len(entries) - 1,
                instruction=instruction,
                toc_tab_pos=toc_tab_pos,
                toc_left_twips=toc_left_twips,
            ))
        settings_root = etree.fromstring(members["word/settings.xml"])
        update_fields = settings_root.find(qn("updateFields"))
        if update_fields is None:
            update_fields = etree.SubElement(settings_root, qn("updateFields"))
        # The TOC now contains page numbers measured from the final render. Do not
        # ask Word to replace this accurate cache with transient zeroes on open.
        # A user can still update the field deliberately after editing the report.
        update_fields.set(qn("val"), "false")
        members["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        members["word/settings.xml"] = etree.tostring(
            settings_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".docx", delete=False) as stream:
            temp = Path(stream.name)
        try:
            with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
                for info in infos:
                    archive.writestr(deepcopy(info), members[info.filename])
            temp.replace(output)
        finally:
            temp.unlink(missing_ok=True)

