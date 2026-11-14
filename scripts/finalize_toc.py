#!/usr/bin/env python3
"""Materialize an accurate cached Word TOC from a rendered report PDF and JSON spec."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from report_toc.finalizer import TocFinalizer


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
    result = TocFinalizer().run(
        source,
        spec,
        args.rendered_pdf.expanduser().resolve(),
        output,
        args.entries_json,
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
