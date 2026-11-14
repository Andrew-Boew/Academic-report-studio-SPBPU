"""Гиперссылки."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class LinkCheck:
    """Закладки и чёрные ссылки без подчёркивания."""

    def apply(self, report) -> None:
        report.bookmarks = {
            node.get(qn("w:name"))
            for node in report.document.element.xpath(".//w:bookmarkStart")
            if node.get(qn("w:name"))
        }
        internal_links = report.document.element.xpath(".//w:hyperlink[@w:anchor]")
        if report.has_toc and not internal_links:
            report.warnings.append("Внутренние гиперссылки содержания ещё не материализованы")
        for link in internal_links:
            anchor = link.get(qn("w:anchor"), "")
            if anchor not in report.bookmarks:
                report.errors.append(f"Внутренняя гиперссылка ведёт на отсутствующую закладку: {anchor}")

        for link in report.document.element.xpath(".//w:hyperlink"):
            for run in link.xpath(".//w:r"):
                if not xml_text(run):
                    continue
                rpr = run.find("./" + qn("w:rPr"))
                color = rpr.find("./" + qn("w:color")) if rpr is not None else None
                underline = rpr.find("./" + qn("w:u")) if rpr is not None else None
                size = rpr.find("./" + qn("w:sz")) if rpr is not None else None
                fonts = rpr.find("./" + qn("w:rFonts")) if rpr is not None else None
                if color is None or color.get(qn("w:val"), "").upper() != "000000":
                    report.errors.append("Гиперссылка должна оставаться чёрной")
                    break
                if underline is None or underline.get(qn("w:val"), "") != "none":
                    report.errors.append("Гиперссылка не должна иметь подчёркивание")
                    break
                if size is None or size.get(qn("w:val")) != "28":
                    report.errors.append("Текст гиперссылки должен иметь размер 14 pt")
                    break
                if fonts is None or fonts.get(qn("w:ascii")) != "Times New Roman" or fonts.get(qn("w:hAnsi")) != "Times New Roman":
                    report.errors.append("Текст гиперссылки должен иметь шрифт Times New Roman")
                    break
