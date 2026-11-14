"""Подписи, формулы и ссылки на них."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class CaptionCheck:
    """Подписи объектов и оформление формул."""

    def apply(self, report) -> None:
        report.figure_numbers: list[str] = []
        report.table_numbers: list[str] = []
        report.listing_numbers: list[str] = []
        report.continuation_numbers: list[str] = []
        report.equation_numbers: list[str] = []
        for paragraph in report.document.paragraphs:
            text_value = paragraph.text.strip()
            if not text_value:
                continue
            figure_match = FIGURE_CAPTION.fullmatch(text_value)
            table_match = TABLE_CAPTION.fullmatch(text_value)
            continuation_match = TABLE_CONTINUATION.fullmatch(text_value)
            listing_match = LISTING_CAPTION.fullmatch(text_value)
            if text_value.lower().startswith("рисунок"):
                if not figure_match:
                    report.errors.append(f"Некорректная подпись рисунка: {text_value!r}")
                else:
                    report.figure_numbers.append(figure_match.group(1).upper())
                if text_value.endswith("."):
                    report.errors.append(f"Точка в конце подписи рисунка: {text_value!r}")
                if effective_alignment(paragraph) != WD_ALIGN_PARAGRAPH.CENTER:
                    report.errors.append(f"Подпись рисунка должна быть по центру: {text_value!r}")
                spacing = paragraph.paragraph_format
                effective_before = spacing.space_before or paragraph.style.paragraph_format.space_before
                effective_after = spacing.space_after or paragraph.style.paragraph_format.space_after
                effective_line = spacing.line_spacing or paragraph.style.paragraph_format.line_spacing
                if effective_before is None or not near(effective_before.pt, 0.0, 0.1):
                    report.errors.append(f"Подпись рисунка должна иметь интервал перед 0 pt: {text_value!r}")
                if effective_after is None or not near(effective_after.pt, 18.0, 0.1):
                    report.errors.append(f"Подпись рисунка должна иметь интервал после 18 pt: {text_value!r}")
                if not isinstance(effective_line, float) or not near(effective_line, report.expected_caption_spacing, 0.01):
                    report.errors.append(
                        f"Подпись рисунка должна иметь межстрочный интервал {report.expected_caption_spacing:g}: {text_value!r}"
                    )
            elif text_value.lower().startswith("таблица"):
                if not table_match:
                    report.errors.append(f"Некорректная подпись таблицы: {text_value!r}")
                else:
                    report.table_numbers.append(table_match.group(1).upper())
                if text_value.endswith("."):
                    report.errors.append(f"Точка в конце подписи таблицы: {text_value!r}")
                expected_table_caption_alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT if report.strict_spbpu else WD_ALIGN_PARAGRAPH.CENTER
                )
                if effective_alignment(paragraph) != expected_table_caption_alignment:
                    alignment_label = "слева" if report.strict_spbpu else "по центру"
                    report.errors.append(f"Подпись таблицы должна быть {alignment_label}: {text_value!r}")
            elif re.match(r"^(?:продолжение|окончание)\s+таблицы", text_value, re.I):
                if not continuation_match:
                    report.errors.append(f"Некорректная подпись переноса таблицы: {text_value!r}")
                else:
                    report.continuation_numbers.append(continuation_match.group(2).upper())
                if effective_alignment(paragraph) != expected_table_caption_alignment:
                    alignment_label = "слева" if report.strict_spbpu else "по центру"
                    report.errors.append(f"Подпись продолжения таблицы должна быть {alignment_label}: {text_value!r}")
            elif text_value.lower().startswith("листинг"):
                if not listing_match:
                    report.errors.append(f"Некорректная подпись листинга: {text_value!r}")
                else:
                    report.listing_numbers.append(listing_match.group(1).upper())
                if text_value.endswith("."):
                    report.errors.append(f"Точка в конце подписи листинга: {text_value!r}")
                if effective_alignment(paragraph) != WD_ALIGN_PARAGRAPH.CENTER:
                    report.errors.append(f"Подпись листинга должна быть по центру: {text_value!r}")
            if paragraph.style and paragraph.style.name.startswith("Report Equation"):
                math_objects = paragraph._p.xpath("./m:oMath")
                if not math_objects:
                    report.errors.append("Формула оформлена обычным текстом, а не Microsoft Word Equation (OMML)")
                elif len(math_objects) != 1:
                    report.errors.append("Один формульный абзац должен содержать ровно один объект Microsoft Word Equation")
                spacing = paragraph.paragraph_format
                effective_before = spacing.space_before or paragraph.style.paragraph_format.space_before
                effective_after = spacing.space_after or paragraph.style.paragraph_format.space_after
                effective_line = spacing.line_spacing or paragraph.style.paragraph_format.line_spacing
                if effective_before is None or not near(effective_before.pt, 8.0, 0.1):
                    report.errors.append("Абзац формулы должен иметь зазор перед 8 pt, отделяющий его от текста")
                if effective_after is None or not near(effective_after.pt, 0.0, 0.1):
                    report.errors.append("Абзац формулы должен иметь интервал после 0 pt")
                if not isinstance(effective_line, float) or not near(effective_line, 1.5, 0.01):
                    report.errors.append("Абзац формулы должен иметь межстрочный интервал 1,5")
                math_sizes = []
                math_text = "".join(node.text or "" for node in paragraph._p.xpath(".//m:t"))
                for math_run in paragraph._p.xpath(".//m:r"):
                    size = math_run.find("./" + qn("w:rPr") + "/" + qn("w:sz"))
                    if size is not None:
                        math_sizes.append(int(size.get(qn("w:val"), "0")))
                longest_number = max((len(value) for value in re.findall(r"\d+", math_text)), default=0)
                expected_size = 28
                if longest_number > 60:
                    report.errors.append(
                        "Целое число длиннее 60 цифр должно быть представлено именованными "
                        "десятичными блоками в отдельных формулах 14 pt"
                    )
                if not math_sizes or any(value != expected_size for value in math_sizes):
                    report.errors.append(
                        "Размер Microsoft Word Equation не соответствует правилу длинного числа: "
                        f"ожидается {expected_size / 2:g} pt"
                    )
                paragraph_size = paragraph._p.find(
                    "./" + qn("w:pPr") + "/" + qn("w:rPr") + "/" + qn("w:sz")
                )
                if paragraph_size is None or paragraph_size.get(qn("w:val")) != str(expected_size):
                    report.errors.append(
                        "Размер формулы должен быть продублирован в свойствах формульного абзаца "
                        "для одинакового рендера в Microsoft Word и LibreOffice"
                    )
                for matrix in paragraph._p.xpath(".//m:m"):
                    rows = matrix.findall("./" + qn("m:mr"))
                    if len(rows) <= 1:
                        continue
                    equation_rows = []
                    for row in rows:
                        row_text = "".join(node.text or "" for node in row.iter(qn("m:t")))
                        cells = row.findall("./" + qn("m:e"))
                        equation_rows.append(len(cells) == 1 and "=" in row_text)
                    if all(equation_rows):
                        report.errors.append(
                            "Самостоятельные равенства или параметры нельзя объединять в одну "
                            "многострочную матрицу; требуется отдельный центрированный абзац для каждой строки"
                        )
                equation_match = re.search(rf"\(({OBJECT_NUMBER})\)\s*$", text_value, re.I)
                if equation_match:
                    tabs = paragraph._p.xpath("./w:pPr/w:tabs/w:tab")
                    tab_values = {tab.get(qn("w:val")) for tab in tabs}
                    if not {"center", "right"}.issubset(tab_values):
                        report.errors.append("Нумерованная формула должна иметь центральную и правую табуляцию")
                elif effective_alignment(paragraph) != WD_ALIGN_PARAGRAPH.CENTER:
                    report.errors.append("Ненумерованная формула или параметр должны быть по центру")
                if equation_match:
                    report.equation_numbers.append(equation_match.group(1).upper())

        duplicate_figures = sorted({number for number in report.figure_numbers if report.figure_numbers.count(number) > 1})
        duplicate_tables = sorted({number for number in report.table_numbers if report.table_numbers.count(number) > 1})
        duplicate_equations = sorted({number for number in report.equation_numbers if report.equation_numbers.count(number) > 1})
        duplicate_listings = sorted({number for number in report.listing_numbers if report.listing_numbers.count(number) > 1})
        if duplicate_figures:
            report.errors.append(f"Повторяются номера рисунков: {', '.join(duplicate_figures)}")
        if duplicate_tables:
            report.errors.append(f"Повторяются номера таблиц: {', '.join(duplicate_tables)}")
        if duplicate_equations:
            report.errors.append(f"Повторяются номера формул: {', '.join(duplicate_equations)}")
        if duplicate_listings:
            report.errors.append(f"Повторяются номера листингов: {', '.join(duplicate_listings)}")
        for number in report.continuation_numbers:
            if number not in report.table_numbers:
                report.errors.append(f"Продолжение ссылается на отсутствующую таблицу {number}")
        warn_numeric_gaps(report.figure_numbers, "рисунков", report.warnings)
        warn_numeric_gaps(report.table_numbers, "таблиц", report.warnings)
        warn_numeric_gaps(report.equation_numbers, "формул", report.warnings)
        warn_numeric_gaps(report.listing_numbers, "листингов", report.warnings)

class ReferenceCheck:
    """Ссылки из текста на рисунки, таблицы, листинги и формулы."""

    def apply(self, report) -> None:
        narrative_parts = []
        for paragraph in report.document.paragraphs:
            value = paragraph.text.strip()
            if not value or FIGURE_CAPTION.fullmatch(value) or TABLE_CAPTION.fullmatch(value) or TABLE_CONTINUATION.fullmatch(value) or LISTING_CAPTION.fullmatch(value):
                continue
            if paragraph.style and paragraph.style.name == "Report Equation":
                continue
            narrative_parts.append(value)
        narrative_text = "\n".join(narrative_parts)
        for number in report.figure_numbers:
            if not re.search(rf"\bрисунк\w*\s+{re.escape(number)}\b", narrative_text, re.I):
                report.errors.append(f"В тексте отсутствует ссылка на рисунок {number}")
            figure_bookmark = figure_bookmark_name(number)
            if figure_bookmark not in report.bookmarks:
                report.errors.append(f"У подписи рисунка {number} отсутствует закладка для перекрёстной ссылки")
        known_equation_numbers = {number.upper() for number in report.equation_numbers}
        for paragraph in report.document.paragraphs:
            value = paragraph.text.strip()
            if (
                not value
                or FIGURE_CAPTION.fullmatch(value)
                or TABLE_CAPTION.fullmatch(value)
                or TABLE_CONTINUATION.fullmatch(value)
                or LISTING_CAPTION.fullmatch(value)
            ):
                continue
            if paragraph.style and paragraph.style.name.startswith("Report Equation"):
                # The "(N)" here is the formula's own label, not a reference.
                continue
            linked_anchors = {
                link.get(qn("w:anchor"), "")
                for link in paragraph._p.xpath(".//w:hyperlink[@w:anchor]")
            }
            for match in FIGURE_REFERENCE.finditer(value):
                figure_bookmark = figure_bookmark_name(match.group(1))
                if figure_bookmark not in linked_anchors:
                    report.errors.append(
                        f"Ссылка на рисунок {match.group(1)} должна быть внутренней гиперссылкой"
                    )
            for match in TABLE_REFERENCE.finditer(value):
                anchor = object_bookmark_name("Table", match.group(1))
                if anchor not in linked_anchors:
                    report.errors.append(
                        f"Ссылка на таблицу {match.group(1)} должна быть внутренней гиперссылкой"
                    )
            for match in LISTING_REFERENCE.finditer(value):
                anchor = object_bookmark_name("Listing", match.group(1))
                if anchor not in linked_anchors:
                    report.errors.append(
                        f"Ссылка на листинг {match.group(1)} должна быть внутренней гиперссылкой"
                    )
            for match in EQUATION_REFERENCE.finditer(value):
                number = match.group(1).upper()
                if number in known_equation_numbers:
                    anchor = object_bookmark_name("Equation", match.group(1))
                    if anchor not in linked_anchors:
                        report.errors.append(
                            f"Ссылка на формулу ({match.group(1)}) должна быть внутренней гиперссылкой"
                        )
            if paragraph.style and paragraph.style.name in {"Normal", "List Bullet", "List Number"}:
                # Plain-text pseudo-formulas are forbidden: every expression with
                # operations must be an inline OMML equation ($...$ in the spec).
                for match in re.finditer(r"[A-Za-z0-9]+\s*[·×]\s*[A-Za-z0-9]+", value):
                    fragment = match.group(0)
                    if any(ch.isascii() and ch.isalpha() for ch in fragment):
                        report.errors.append(
                            f"Выражение {fragment!r} набрано обычным текстом; оформите его "
                            "inline-формулой $…$ (m:oMath), а не буквами со знаком умножения"
                        )
                for match in re.finditer(r"(?<![A-Za-z0-9])[A-Za-z]\s*/\s*[A-Za-z0-9](?![A-Za-z0-9])", value):
                    report.errors.append(
                        f"Дробь {match.group(0)!r} записана через слэш; оформите её "
                        "inline-формулой $…$ с настоящей дробью OMML"
                    )
                for match in re.finditer(r"\b[A-Za-z]\s*=\s*\S", value):
                    report.errors.append(
                        f"Равенство {match.group(0)!r} набрано обычным текстом; оформите его "
                        "inline-формулой $…$ (m:oMath)"
                    )
        for number in report.table_numbers:
            if not re.search(rf"\bтабл(?:иц\w*|\.)\s+{re.escape(number)}\b", narrative_text, re.I):
                report.errors.append(f"В тексте отсутствует ссылка на таблицу {number}")
        for number in report.listing_numbers:
            if not re.search(rf"\bлистинг\w*\s+{re.escape(number)}\b", narrative_text, re.I):
                report.errors.append(f"В тексте отсутствует ссылка на листинг {number}")
        for number in report.equation_numbers:
            if not re.search(rf"(?:формул\w*|уравнен\w*)\s*\({re.escape(number)}\)", narrative_text, re.I):
                report.warnings.append(f"В тексте не найдена явная ссылка на формулу ({number})")
