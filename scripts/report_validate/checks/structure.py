"""Заголовки, реферат и список источников."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class HeadingCheck:
    """Текст заголовков и разрыв страницы у раздела."""

    def apply(self, report) -> None:
        heading_style_names = {"Heading 1", "Heading 2", "Heading 3", "Report Structural Heading"}
        report.heading_texts = [
            p.text.strip() for p in report.document.paragraphs
            if p.style and p.style.name in heading_style_names
        ]
        if not report.heading_texts:
            report.errors.append("Не найдены настоящие стили заголовков Word")
        for paragraph in report.document.paragraphs:
            if not paragraph.style or paragraph.style.name not in heading_style_names:
                continue
            value = paragraph.text.strip()
            if paragraph.style.name.startswith("Heading "):
                if MANUAL_HEADING_NUMBER.match(value):
                    report.errors.append(f"Номер главы набран вручную вместо функции нумерации Word: {value!r}")
                if paragraph.style.name == "Heading 1" and not paragraph._p.xpath("./w:pPr/w:pageBreakBefore"):
                    report.errors.append(
                        f"Каждый раздел основной части должен начинаться с новой страницы: {value!r}"
                    )
                direct_left = paragraph.paragraph_format.left_indent
                direct_first = paragraph.paragraph_format.first_line_indent
                left_cm = direct_left.cm if direct_left is not None else 0.0
                first_cm = direct_first.cm if direct_first is not None else 0.0
                if not near(left_cm, 0.0) or not near(first_cm, 1.25):
                    report.errors.append(
                        f"{paragraph.style.name}: номер заголовка должен начинаться на 1,25 см "
                        f"(прямой левый отступ 0 см, первая строка 1,25 см)"
                    )
            words = re.findall(r"[А-ЯЁа-яё]{3,}", value)
            if paragraph.style.name == "Report Structural Heading" and report.strict_spbpu:
                if words and not all(word.isupper() for word in words):
                    report.errors.append(f"Структурный заголовок должен быть прописными: {value!r}")
            elif words and len(words) >= 2 and all(word.isupper() for word in words):
                report.errors.append(f"Заголовок должен быть в регистре предложения, а не прописными: {value!r}")

class FrontMatterCheck:
    """Реферат, перечень сокращений и порядок приложений."""

    def apply(self, report) -> None:
        abstract_heading_present = any(
            paragraph.style and paragraph.style.name == "Report Structural Heading"
            and paragraph.text.strip().upper() == "РЕФЕРАТ"
            for paragraph in report.document.paragraphs
        )
        if abstract_heading_present:
            keywords_line = next(
                (p.text for p in report.document.paragraphs if p.text.strip().upper().startswith("КЛЮЧЕВЫЕ СЛОВА")),
                "",
            )
            keyword_words = [
                word.strip() for word in keywords_line.split(":", 1)[-1].split(",") if word.strip()
            ] if keywords_line else []
            if not (5 <= len(keyword_words) <= 15):
                report.errors.append("РЕФЕРАТ: ключевых слов должно быть от 5 до 15 (tr23-26 2.3.2)")
            if not any(
                paragraph.style and paragraph.style.name == "Report Structural Heading"
                and paragraph.text.strip().upper() == "ABSTRACT"
                for paragraph in report.document.paragraphs
            ):
                report.errors.append("РЕФЕРАТ должен сопровождаться английской версией ABSTRACT (tr23-26 2.3.4)")
            if "Работа содержит" not in report.all_text:
                report.errors.append("РЕФЕРАТ должен начинаться со сведений об объёме работы («Работа содержит: …»)")
        abbreviations_present = any(
            paragraph.text.strip().upper() == "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ"
            for paragraph in report.document.paragraphs
        )
        if abbreviations_present and "В настоящей работе применяют" not in report.all_text:
            report.errors.append(
                "Перечень сокращений должен начинаться со слов «В настоящей работе применяют "
                "следующие сокращения и обозначения» (tr23-26 2.4.1)"
            )

        appendix_indices = [i for i, text in enumerate(report.heading_texts) if text.upper().startswith("ПРИЛОЖЕНИЕ")]
        source_indices = [i for i, text in enumerate(report.heading_texts) if "ИСТОЧНИК" in text.upper() or "ЛИТЕРАТУР" in text.upper()]
        if appendix_indices and source_indices and min(appendix_indices) < max(source_indices):
            report.errors.append("Приложения должны располагаться после списка источников")

class SourceCheck:
    """Источники, внешние ссылки, перечисления и цитаты."""

    def apply(self, report) -> None:
        report.bibliography_paragraphs = [
            paragraph for paragraph in report.document.paragraphs
            if paragraph.style and paragraph.style.name == "Report Bibliography"
        ]
        if report.cover_layout != "lab" and not report.bibliography_paragraphs:
            report.errors.append("Курсовая работа, ВКР и отчёт по НИР должны содержать список проверенных реальных источников")
        if report.cover_layout == "lab" and report.bibliography_paragraphs:
            report.errors.append("Лабораторная работа не должна содержать список использованных источников")
        report.seen_external_links = 0
        for paragraph in report.document.paragraphs:
            external_links = paragraph._p.xpath(".//w:hyperlink[@r:id]")
            if external_links and (not paragraph.style or paragraph.style.name != "Report Bibliography"):
                report.errors.append("Внешние гиперссылки разрешены только в списке использованных источников")
            if paragraph.style and paragraph.style.name == "Report Bibliography":
                if len(external_links) != 1:
                    report.errors.append("Каждый источник должен содержать ровно одну проверенную внешнюю гиперссылку")
            for link in external_links:
                report.seen_external_links += 1
                relation_id = link.get(qn("r:id"), "")
                relation = paragraph.part.rels.get(relation_id)
                target = relation.target_ref if relation is not None else ""
                parsed = urlparse(target)
                host = (parsed.hostname or "").casefold()
                if parsed.scheme != "https" or not host or host in {"localhost", "127.0.0.1", "::1"}:
                    report.errors.append(f"Источник должен вести на публичный HTTPS-адрес: {target!r}")
        all_external_links = report.document.element.xpath(".//w:hyperlink[@r:id]")
        if report.seen_external_links != len(all_external_links):
            report.errors.append("Обнаружена внешняя гиперссылка вне обычного элемента списка источников")
        if report.bibliography_paragraphs and report.seen_external_links != len(report.bibliography_paragraphs):
            report.errors.append("Число гиперссылок не совпадает с числом источников")
        if report.report_type == "lab" and report.seen_external_links:
            report.errors.append("В лабораторной работе запрещены внешние гиперссылки")

        narrative_citations: list[int] = []
        for paragraph in report.document.paragraphs:
            if not paragraph.style or paragraph.style.name not in {"Normal", "List Bullet", "List Number"}:
                continue
            narrative_citations.extend(int(value) for value in re.findall(r"\[(\d+)\]", paragraph.text))
        if report.report_type == "lab" and narrative_citations:
            report.errors.append("В лабораторной работе не должно быть ссылок на литературу вида [N]")
        if report.strict_spbpu:
            # SPbPU tr23-26 3.5.2: a simple enumeration level is a dash, not a bullet.
            bullet_numbering = report.document.part.numbering_part.element
            list_bullet_style = report.document.styles["List Bullet"]
            bullet_num_pr = list_bullet_style._element.get_or_add_pPr().find(qn("w:numPr"))
            bullet_marker_ok = False
            if bullet_num_pr is not None:
                bullet_num_id_node = bullet_num_pr.find(qn("w:numId"))
                if bullet_num_id_node is not None:
                    bullet_num_id = bullet_num_id_node.get(qn("w:val"))
                    bullet_num = next(
                        (node for node in bullet_numbering.findall(qn("w:num")) if node.get(qn("w:numId")) == bullet_num_id),
                        None,
                    )
                    abstract_ref = bullet_num.find(qn("w:abstractNumId")) if bullet_num is not None else None
                    abstract = next(
                        (
                            node for node in bullet_numbering.findall(qn("w:abstractNum"))
                            if node.get(qn("w:abstractNumId")) == (abstract_ref.get(qn("w:val")) if abstract_ref is not None else None)
                        ),
                        None,
                    )
                    lvl_text = abstract.find(".//" + qn("w:lvl") + "/" + qn("w:lvlText")) if abstract is not None else None
                    bullet_marker_ok = lvl_text is not None and lvl_text.get(qn("w:val")) == "–"
            if not bullet_marker_ok:
                report.errors.append(
                    "Первый уровень перечисления должен использовать тире (Word-нумерация с маркером «–»)"
                )
        if report.cover_layout != "lab":
            for number in narrative_citations:
                if number < 1 or number > len(report.bibliography_paragraphs):
                    report.errors.append(f"Ссылка [{number}] не соответствует элементу списка источников")
            for link in report.document.element.xpath(".//w:hyperlink[starts-with(@w:anchor, '_ReportBib')]"):
                anchor = link.get(qn("w:anchor"), "")
                if anchor not in report.bookmarks:
                    report.errors.append(f"Ссылка на источник ведёт на отсутствующую закладку: {anchor}")

class ProseCheck:
    """Отступ и выравнивание обычного текста."""

    def apply(self, report) -> None:
        ordinary = [
            p for p in report.document.paragraphs
            if p.text.strip() and p.style and p.style.name == "Normal"
        ]
        bad_indents = 0
        for paragraph in ordinary:
            direct = paragraph.paragraph_format.first_line_indent
            if (
                direct is not None
                and not near(direct.cm, 1.25)
                and not (near(direct.cm, 0.0) and effective_alignment(paragraph) in {WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT})
            ):
                bad_indents += 1
        if bad_indents:
            report.warnings.append(f"Обычных абзацев с прямым отступом не 1,25 см: {bad_indents}")
        left_aligned_body = [
            p for p in ordinary
            if effective_alignment(p) == WD_ALIGN_PARAGRAPH.LEFT
        ]
        if left_aligned_body:
            report.errors.append(
                "Обычный текст должен быть выровнен по ширине; абзацев с выравниванием влево: "
                f"{len(left_aligned_body)}"
            )
