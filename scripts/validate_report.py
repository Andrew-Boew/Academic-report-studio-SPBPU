#!/usr/bin/env python3
"""Validate structural and formatting invariants of a generated academic DOCX."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from report_validate.inspector import inspect


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    args = parser.parse_args()
    result = inspect(args.docx.expanduser().resolve())
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(payload, encoding="utf-8")
    print(payload)
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
