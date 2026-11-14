"""Объекты в потоке документа."""

from __future__ import annotations

import re
import zipfile
from urllib.parse import urlparse

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from report_validate.support import *

class BodyObjectCheck:
    """Рисунки, рамки кода и таблицы относительно подписей."""

    def apply(self, report) -> None:
        body_items = list(report.document.element.body.iterchildren())
        seen_first_page_break = False
        report.content_tables = 0
        for item_index, item in enumerate(body_items):
            if item.tag == qn("w:p"):
                page_break = any(
                    node.get(qn("w:type")) == "page"
                    for node in item.findall(".//" + qn("w:br"))
                )
                if page_break:
                    seen_first_page_break = True
                has_drawing = item.find(".//" + qn("w:drawing")) is not None or item.find(".//" + qn("w:pict")) is not None
                if has_drawing and seen_first_page_break:
                    next_text = ""
                    for following in body_items[item_index + 1:]:
                        if following.tag == qn("w:p") and not xml_text(following):
                            continue
                        next_text = xml_text(following) if following.tag == qn("w:p") else "<table>"
                        break
                    if not FIGURE_CAPTION.fullmatch(next_text):
                        report.errors.append("Подпись рисунка должна находиться непосредственно под изображением")
            elif item.tag == qn("w:tbl") and not is_layout_table(item):
                if is_figure_placeholder(item):
                    next_text = ""
                    for following in body_items[item_index + 1:]:
                        if following.tag == qn("w:p") and not xml_text(following):
                            continue
                        next_text = xml_text(following) if following.tag == qn("w:p") else "<table>"
                        break
                    if not FIGURE_CAPTION.fullmatch(next_text):
                        report.errors.append("Подпись места под рисунок должна находиться непосредственно под рамкой")
                    continue
                if is_code_fragment(item):
                    previous_text = ""
                    for previous in reversed(body_items[:item_index]):
                        if previous.tag == qn("w:p") and not xml_text(previous):
                            continue
                        previous_text = xml_text(previous) if previous.tag == qn("w:p") else "<table>"
                        break
                    if not LISTING_CAPTION.fullmatch(previous_text):
                        report.errors.append("Рамка короткого кода должна иметь подпись `Листинг N – Название` непосредственно сверху")
                    rows = item.findall("./" + qn("w:tr"))
                    cells = item.findall("./" + qn("w:tr") + "/" + qn("w:tc"))
                    if len(rows) != 1 or len(cells) != 1:
                        report.errors.append("Короткий фрагмент кода должен находиться в одноячеечной рамке")
                    for run in item.findall(".//" + qn("w:r")):
                        if not xml_text(run):
                            continue
                        rpr = run.find("./" + qn("w:rPr"))
                        fonts = rpr.find("./" + qn("w:rFonts")) if rpr is not None else None
                        size = rpr.find("./" + qn("w:sz")) if rpr is not None else None
                        font_name = fonts.get(qn("w:ascii")) if fonts is not None else None
                        half_points = int(size.get(qn("w:val"), "0")) if size is not None else 0
                        if font_name != "Courier New" or half_points not in {22, 24}:
                            report.errors.append("Короткий код в рамке должен быть Courier New 11 или 12 pt")
                            break
                    continue
                report.content_tables += 1
                previous_text = ""
                for previous in reversed(body_items[:item_index]):
                    if previous.tag == qn("w:p") and not xml_text(previous):
                        continue
                    previous_text = xml_text(previous) if previous.tag == qn("w:p") else "<table>"
                    break
                if not (TABLE_CAPTION.fullmatch(previous_text) or TABLE_CONTINUATION.fullmatch(previous_text)):
                    report.errors.append(f"Таблица {report.content_tables}: подпись должна находиться непосредственно сверху")

                table_rows = item.findall("./" + qn("w:tr"))
                if not table_rows:
                    report.errors.append(f"Таблица {report.content_tables}: отсутствуют строки")
                    continue
                header = table_rows[0].find("./" + qn("w:trPr") + "/" + qn("w:tblHeader"))
                if header is None or header.get(qn("w:val"), "true") not in {"true", "1", "on"}:
                    report.errors.append(f"Таблица {report.content_tables}: строка заголовка не отмечена для повтора")
                if len(table_rows) > 1:
                    for header_paragraph in table_rows[0].findall(".//" + qn("w:p")):
                        keep_next = header_paragraph.find("./" + qn("w:pPr") + "/" + qn("w:keepNext"))
                        if keep_next is None:
                            report.errors.append(
                                f"Таблица {report.content_tables}: шапка должна удерживаться минимум "
                                "с первой строкой данных"
                            )
                            break
                table_style = item.find("./" + qn("w:tblPr") + "/" + qn("w:tblStyle"))
                if table_style is None or table_style.get(qn("w:val")) not in {"TableGrid", "Table Grid"}:
                    report.errors.append(f"Таблица {report.content_tables}: требуются тонкие внешние и внутренние границы")
                small_runs = []
                for row_index, row in enumerate(table_rows, 1):
                    for cell_index, cell in enumerate(row.findall("./" + qn("w:tc")), 1):
                        no_wrap = cell.find("./" + qn("w:tcPr") + "/" + qn("w:noWrap"))
                        fit_text = cell.find("./" + qn("w:tcPr") + "/" + qn("w:tcFitText"))
                        if no_wrap is not None or fit_text is not None:
                            report.errors.append(
                                f"Таблица {report.content_tables}: w:noWrap/w:tcFitText запрещены; "
                                "длинное значение должно переноситься естественно"
                            )
                        cell_text = xml_text(cell).strip().replace("−", "-")
                        if re.fullmatch(r"[+-]?\d+(?:[.,]\d+)?", cell_text):
                            for cell_paragraph in cell.findall("./" + qn("w:p")):
                                alignment = cell_paragraph.find(
                                    "./" + qn("w:pPr") + "/" + qn("w:jc")
                                )
                                if alignment is None or alignment.get(qn("w:val")) != "center":
                                    report.errors.append(
                                        f"Таблица {report.content_tables}: числовая ячейка "
                                        f"{row_index}:{cell_index} должна быть по центру"
                                    )
                                    break
                        for run in cell.findall(".//" + qn("w:r")):
                            if not xml_text(run):
                                continue
                            rpr = run.find("./" + qn("w:rPr"))
                            size = rpr.find("./" + qn("w:sz")) if rpr is not None else None
                            fonts = rpr.find("./" + qn("w:rFonts")) if rpr is not None else None
                            if size is not None:
                                half_points = int(size.get(qn("w:val"), "0"))
                                font_name = fonts.get(qn("w:ascii")) if fonts is not None else None
                                if half_points < round(report.expected_table_size_pt * 2):
                                    small_runs.append(f"{row_index}:{cell_index}={half_points / 2:g}pt")
                            if row_index == 1:
                                bold = run.find("./" + qn("w:rPr") + "/" + qn("w:b"))
                                if bold is not None and bold.get(qn("w:val"), "true") in {"true", "1", "on"}:
                                    report.errors.append(
                                        f"Таблица {report.content_tables}: полужирный в таблице не применяется "
                                        "(ГОСТ 7.32-2017 6.1.1 — только для заголовков разделов)"
                                    )
                                    break
                if small_runs:
                    report.errors.append(
                        f"Таблица {report.content_tables}: табличный текст меньше обязательных {report.expected_table_size_pt:g} pt: "
                        + ", ".join(small_runs[:10])
                    )
                shaded_cells = []
                for row_index, row in enumerate(table_rows, 1):
                    for cell_index, cell in enumerate(row.findall("./" + qn("w:tc")), 1):
                        shading = cell.find("./" + qn("w:tcPr") + "/" + qn("w:shd"))
                        if shading is None:
                            continue
                        fill = shading.get(qn("w:fill"), "").upper()
                        if fill not in {"", "AUTO", "FFFFFF", "CLEAR"}:
                            shaded_cells.append(f"{row_index}:{cell_index}={fill}")
                if shaded_cells:
                    report.errors.append(
                        f"Таблица {report.content_tables}: заливка ячеек запрещена локальным профилем: "
                        + ", ".join(shaded_cells[:10])
                    )
                split_rows = []
                for row_index, row in enumerate(table_rows, 1):
                    cant_split = row.find("./" + qn("w:trPr") + "/" + qn("w:cantSplit"))
                    if cant_split is None or cant_split.get(qn("w:val"), "true") not in {"true", "1", "on"}:
                        split_rows.append(row_index)
                if split_rows:
                    report.errors.append(
                        f"Таблица {report.content_tables}: Word может разорвать строки между страницами: "
                        + ", ".join(map(str, split_rows[:10]))
                    )
