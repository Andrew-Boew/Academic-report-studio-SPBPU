"""Преобразование LaTeX в формулу Word (OMML)."""

from __future__ import annotations

import copy
import datetime as dt
import functools
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlparse

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Cm, Pt, RGBColor
from lxml import etree

from report_build.constants import *
from report_build.errors import SpecError

def style_math_runs(node, half_points: str = "28") -> None:
    """Force Cambria Math and the given size on every math run of the node."""
    for math_run in node.xpath(".//m:r", namespaces=OMML_NAMESPACES):
        rpr = math_run.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            math_run.insert(0, rpr)
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rpr.insert(0, fonts)
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(f"w:{key}"), "Cambria Math")
        for tag in ("w:sz", "w:szCs"):
            size = rpr.find(qn(tag))
            if size is None:
                size = OxmlElement(tag)
                rpr.append(size)
            size.set(qn("w:val"), half_points)
        character_spacing = rpr.find(qn("w:spacing"))
        if character_spacing is not None:
            rpr.remove(character_spacing)

@functools.lru_cache(maxsize=256)
def latex_to_omath(latex: str):
    pandoc = shutil.which("pandoc")
    if not pandoc:
        raise SpecError("Pandoc is required to create native Microsoft Word equations")
    source = f"$$\n{latex}\n$$\n"
    with tempfile.TemporaryDirectory(prefix="report-equation-") as tmp:
        output = Path(tmp) / "equation.docx"
        process = subprocess.run(
            [pandoc, "--from", "markdown", "--to", "docx", "--output", str(output)],
            input=source,
            text=True,
            capture_output=True,
        )
        if process.returncode != 0 or not output.is_file():
            message = process.stderr.strip() or "unknown Pandoc error"
            raise SpecError(f"Cannot convert LaTeX equation to OMML: {message}")
        with zipfile.ZipFile(output) as archive:
            xml = archive.read("word/document.xml")
    root = etree.fromstring(xml)
    nodes = root.xpath(".//m:oMath", namespaces=OMML_NAMESPACES)
    if not nodes:
        raise SpecError("Pandoc produced no OMML equation")
    return copy.deepcopy(nodes[0])
