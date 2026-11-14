"""Служебные части пакета DOCX."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class PackageCheck:
    """Заполнители, сноски и признаки полей Word."""

    def apply(self, report) -> None:
        report.all_text = "\n".join(p.text for p in report.document.paragraphs)
        for table in report.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    report.all_text += "\n" + cell.text
        if PLACEHOLDER.search(report.all_text):
            report.errors.append("В документе остались служебные заполнители")
        if re.search(r"\d[ \u00a0]+°", report.all_text):
            report.errors.append("Перед поднятым знаком градуса ° не должно быть пробела: 120° (tr23-26 3.10.1)")

        with zipfile.ZipFile(report.path) as archive:
            document_xml = archive.read("word/document.xml")
            settings_xml = archive.read("word/settings.xml")
            footer_xml = b"".join(
                archive.read(name) for name in archive.namelist() if name.startswith("word/footer") and name.endswith(".xml")
            )
            names = archive.namelist()
            footnotes_xml = archive.read("word/footnotes.xml") if "word/footnotes.xml" in names else b""
        referenced_footnotes = set(re.findall(rb'footnoteReference[^>]*w:id="(\d+)"', document_xml))
        defined_footnotes = set(re.findall(rb'<w:footnote[^>]*w:id="(-?\d+)"', footnotes_xml))
        if referenced_footnotes and not footnotes_xml:
            report.errors.append("В тексте есть сноски, но часть word/footnotes.xml отсутствует")
        if referenced_footnotes - defined_footnotes:
            report.errors.append("Сноска в тексте ссылается на отсутствующую запись footnotes.xml")
        if footnotes_xml and not referenced_footnotes:
            report.warnings.append("Часть footnotes.xml есть, но в тексте нет ни одной ссылки на сноску")
        report.has_toc = b"TOC " in document_xml
        if not report.has_toc:
            report.warnings.append("Поле содержания TOC не найдено")
        if b"updateFields" not in settings_xml:
            report.errors.append("Не установлен автоматический запрос на обновление полей Word")
        if b"PAGE " not in footer_xml:
            report.errors.append("Поле номера страницы PAGE не найдено в нижнем колонтитуле")
        if "Обновите содержание в Word".encode("utf-8") in document_xml:
            report.warnings.append("После открытия в Word нужно обновить содержание (Ctrl+A, F9)")
