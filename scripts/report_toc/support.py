"""Текст, страницы PDF и узлы Word для содержания. Без сценария сборки."""

from __future__ import annotations

import re
from lxml import etree

NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": NS_URI}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

def qn(name: str) -> str:
    return f"{{{NS_URI}}}{name}"

def normalized(value: str) -> str:
    value = value.replace("\u00a0", " ").replace("—", "–")
    return re.sub(r"\s+", " ", value).strip().casefold()

def structural_heading(block: dict) -> bool:
    if "structural" in block:
        return bool(block["structural"])
    text = str(block.get("text", "")).strip()
    return int(block.get("level", 1)) == 1 and bool(re.match(
        r"^(?:СОДЕРЖАНИЕ|ВВЕДЕНИЕ|ЗАКЛЮЧЕНИЕ|ВЫВОДЫ|СПИСОК|ПРИЛОЖЕНИЕ)", text, re.I
    ))

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

def entry_paragraph(
    entry: dict,
    *,
    first: bool,
    last: bool,
    instruction: str,
    toc_tab_pos: str = "9354",
    toc_left_twips: dict[int, str] | None = None,
) -> etree._Element:
    paragraph = etree.Element(qn("p"))
    ppr = etree.SubElement(paragraph, qn("pPr"))
    style = etree.SubElement(ppr, qn("pStyle"))
    style.set(qn("val"), f"TOC{entry['level']}")
    tabs = etree.SubElement(ppr, qn("tabs"))
    tab = etree.SubElement(tabs, qn("tab"))
    tab.set(qn("val"), "right")
    tab.set(qn("leader"), "dot")
    # The tab sits exactly at the working text width so page numbers reach the
    # right margin: 16.5 cm (9354 twips) for student_default and 17.0 cm
    # (9638 twips) for spbpu. Keeping the tab inside that boundary
    # prevents Word from clipping or dropping page numbers that LibreOffice
    # may still render outside it.
    tab.set(qn("pos"), toc_tab_pos)
    # Write the 1.5 line interval explicitly so the materialized contents keep
    # it even when the document TOC styles are missing or stale.
    spacing = etree.SubElement(ppr, qn("spacing"))
    spacing.set(qn("before"), "0")
    spacing.set(qn("after"), "0")
    spacing.set(qn("line"), "360")
    spacing.set(qn("lineRule"), "auto")
    indent = etree.SubElement(ppr, qn("ind"))
    # Contents hierarchy is expressed by paragraph indents, not by spaces:
    # TOC 1 = 0 cm always; deeper levels follow the active profile
    # (student_default: 1.25/2.50 cm, spbpu: 0.5/1.0 cm).
    if toc_left_twips is None:
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

def fill_abstract_volume(root: etree._Element, page_count: int) -> int:
    """Replace the `__` page-count placeholders in РЕФЕРАТ/ABSTRACT volume lines."""
    filled = 0
    for paragraph in root.xpath(".//w:body//w:p", namespaces=NS):
        text = paragraph_text(paragraph)
        if not (text.startswith("Работа содержит") or text.startswith("The work contains")):
            continue
        if not re.search(r"_{2,}", text):
            continue
        for node in paragraph.xpath(".//w:t", namespaces=NS):
            value = node.text or ""
            if re.search(r"_{2,}", value):
                node.text = re.sub(r"_{2,}", str(page_count), value)
                filled += 1
    return filled
