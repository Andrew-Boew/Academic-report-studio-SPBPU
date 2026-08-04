#!/usr/bin/env python3
"""Materialize an accurate cached Word TOC from a rendered report PDF and JSON spec."""

from __future__ import annotations

import argparse
import json
import re
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree
from pypdf import PdfReader


NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": NS_URI}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def qn(name: str) -> str:
    return f"{{{NS_URI}}}{name}"


def normalized(value: str) -> str:
    value = value.replace("\u00a0", " ").replace("–", "—")
    return re.sub(r"\s+", " ", value).strip().casefold()


def structural_heading(block: dict) -> bool:
    if "structural" in block:
        return bool(block["structural"])
    text = str(block.get("text", "")).strip()
    return int(block.get("level", 1)) == 1 and bool(re.match(
        r"^(?:СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ВЫВОДЫ|СПИСОК|ПРИЛОЖЕНИЕ)", text, re.I
    ))


def collect_entries(spec: dict) -> list[dict]:
    levels = max(1, min(3, int((spec.get("document") or {}).get("toc_levels", 3))))
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
                    display_text = f"{'.'.join(str(value) for value in counters[:level])}. {text}"
                entries.append({"text": text, "display_text": display_text, "level": level})
        elif kind == "bibliography":
            text = str(block.get("title", "Список использованных источников")).strip()
            entries.append({"text": text, "display_text": text, "level": 1})
        elif kind == "appendix_heading":
            text = f"Приложение {str(block.get('letter', '')).strip().upper()}"
            entries.append({"text": text, "display_text": text, "level": 1})
    return entries


def squashed(value: str) -> str:
    """Normalized text without any whitespace.

    Justified lines in Word PDFs can be extracted with stretched gaps inside
    words (``Python      -обёртки``); removing all whitespace makes the
    comparison immune to that while dot leaders and page numbers still keep
    TOC lines distinct from real headings.
    """
    return re.sub(r"\s+", "", normalized(value))


def page_lines(page) -> list[list[str]]:
    """Return normalized text lines in plain and layout extraction modes.

    LibreOffice PDFs keep paragraph line breaks in the default mode, while
    Word PDFs may join a whole page into one line there; the layout mode of
    pypdf restores per-line text for such files.
    """
    variants = []
    plain = (page.extract_text() or "").splitlines()
    variants.append([normalized(line) for line in plain if line.strip()])
    try:
        layout = page.extract_text(extraction_mode="layout") or ""
    except Exception:
        layout = ""
    variants.append([normalized(line) for line in layout.splitlines() if line.strip()])
    return variants


def locate_pages(entries: list[dict], pdf_path: Path, page_start: int) -> list[dict]:
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


def run_with_text(text: str, *, bold: bool = False) -> etree._Element:
    run = etree.Element(qn("r"))
    rpr = etree.SubElement(run, qn("rPr"))
    fonts = etree.SubElement(rpr, qn("rFonts"))
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        fonts.set(qn(key), "Times New Roman")
    color = etree.SubElement(rpr, qn("color"))
    color.set(qn("val"), "000000")
    underline = etree.SubElement(rpr, qn("u"))
    underline.set(qn("val"), "none")
    size = etree.SubElement(rpr, qn("sz"))
    size.set(qn("val"), "28")
    size_cs = etree.SubElement(rpr, qn("szCs"))
    size_cs.set(qn("val"), "28")
    if bold:
        etree.SubElement(rpr, qn("b"))
    node = etree.SubElement(run, qn("t"))
    node.text = text
    return run


def field_run(field_type: str) -> etree._Element:
    run = etree.Element(qn("r"))
    node = etree.SubElement(run, qn("fldChar"))
    node.set(qn("fldCharType"), field_type)
    if field_type == "begin":
        node.set(qn("dirty"), "false")
    return run


def entry_paragraph(entry: dict, *, first: bool, last: bool, instruction: str) -> etree._Element:
    paragraph = etree.Element(qn("p"))
    ppr = etree.SubElement(paragraph, qn("pPr"))
    style = etree.SubElement(ppr, qn("pStyle"))
    style.set(qn("val"), f"TOC{entry['level']}")
    tabs = etree.SubElement(ppr, qn("tabs"))
    tab = etree.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "right")
    tab.set(qn("leader"), "dot")
    # A4 text width for 3.0 cm left and 1.5 cm right margins: 16.5 cm.
    # Keeping the tab inside that boundary prevents Word from clipping or
    # dropping page numbers that LibreOffice may still render outside it.
    tab.set(qn("pos"), "9354")
    # Write the 1.5 line interval explicitly so the materialized contents keep
    # it even when the document TOC styles are missing or stale.
    spacing = etree.SubElement(ppr, qn("spacing"))
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    indent = etree.SubElement(ppr, qn("ind"))
    # Contents hierarchy is expressed by paragraph indents, not by spaces:
    # TOC 1 = 0 cm, TOC 2 = 1.25 cm, TOC 3 = 2.50 cm.
    toc_left_twips = {1: "0", 2: "709", 3: "1417"}
    indent.set(qn("left"), toc_left_twips[entry["level"]])
    indent.set(qn("firstLine"), "0")
    if first:
        paragraph.append(field_run("begin"))
        instruction_run = etree.SubElement(paragraph, qn("r"))
        instruction_node = etree.SubElement(instruction_run, qn("instrText"))
        instruction_node.set(XML_SPACE, "preserve")
        instruction_node.text = instruction
        paragraph.append(field_run("separate"))
    link = etree.SubElement(paragraph, qn("hyperlink"))
    link.set(qn("anchor"), entry["anchor"])
    link.set(qn("history"), "1")
    link.append(run_with_text(entry.get("display_text", entry["text"])))
    tab_run = etree.SubElement(link, qn("r"))
    tab_rpr = etree.SubElement(tab_run, qn("rPr"))
    tab_fonts = etree.SubElement(tab_rpr, qn("rFonts"))
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        tab_fonts.set(qn(key), "Times New Roman")
    tab_color = etree.SubElement(tab_rpr, qn("color"))
    tab_color.set(qn("val"), "000000")
    tab_underline = etree.SubElement(tab_rpr, qn("u"))
    tab_underline.set(qn("val"), "none")
    tab_size = etree.SubElement(tab_rpr, qn("sz"))
    tab_size.set(qn("val"), "28")
    tab_size_cs = etree.SubElement(tab_rpr, qn("szCs"))
    tab_size_cs.set(qn("val"), "28")
    etree.SubElement(tab_run, qn("tab"))
    link.append(run_with_text(str(entry["page"])))
    if last:
        paragraph.append(field_run("end"))
    return paragraph


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def attach_heading_bookmarks(root: etree._Element, entries: list[dict]) -> None:
    used_ids = []
    for node in root.xpath(".//w:bookmarkStart", namespaces=NS):
        try:
            used_ids.append(int(node.get(qn("id"), "0")))
        except ValueError:
            continue
    next_id = max(used_ids, default=0) + 1
    paragraphs = root.xpath(".//w:body//w:p", namespaces=NS)
    for index, entry in enumerate(entries, start=1):
        target = normalized(entry["text"])
        candidates = [paragraph for paragraph in paragraphs if normalized(paragraph_text(paragraph)) == target]
        if not candidates:
            raise ValueError(f"Heading paragraph not found in DOCX: {entry['text']}")
        heading = candidates[-1]
        existing = heading.xpath("./w:bookmarkStart[starts-with(@w:name, '_ReportToc')]", namespaces=NS)
        if existing:
            entry["anchor"] = existing[0].get(qn("name"))
            continue
        anchor = f"_ReportToc{index:04d}"
        bookmark_start = etree.Element(qn("bookmarkStart"))
        bookmark_start.set(qn("id"), str(next_id))
        bookmark_start.set(qn("name"), anchor)
        bookmark_end = etree.Element(qn("bookmarkEnd"))
        bookmark_end.set(qn("id"), str(next_id))
        ppr = heading.find(qn("pPr"))
        heading.insert(1 if ppr is not None else 0, bookmark_start)
        heading.append(bookmark_end)
        entry["anchor"] = anchor
        next_id += 1


def patch_docx(source: Path, output: Path, entries: list[dict]) -> None:
    with zipfile.ZipFile(source) as archive:
        infos = archive.infolist()
        members = {info.filename: archive.read(info.filename) for info in infos}
    root = etree.fromstring(members["word/document.xml"])
    attach_heading_bookmarks(root, entries)
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
    for offset, entry in enumerate(entries):
        parent.insert(index + offset, entry_paragraph(
            entry,
            first=offset == 0,
            last=offset == len(entries) - 1,
            instruction=instruction,
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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--spec", required=True, type=Path)
    parser.add_argument("--rendered-pdf", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--entries-json", type=Path)
    args = parser.parse_args()
    source = args.docx.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if source == output:
        raise SystemExit("Use a distinct --output path")
    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    page_start = int((spec.get("document") or {}).get("page_number_start", 0))
    entries = locate_pages(collect_entries(spec), args.rendered_pdf.expanduser().resolve(), page_start)
    if not entries:
        raise SystemExit("No TOC entries found in the specification")
    patch_docx(source, output, entries)
    if args.entries_json:
        args.entries_json.parent.mkdir(parents=True, exist_ok=True)
        args.entries_json.write_text(json.dumps({"entries": entries}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": str(output), "entries": entries}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
