#!/usr/bin/env python3
"""Create a privacy-conscious inventory of files used for an academic report."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from report_inventory.inventory import Inventory


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("--out", required=True, type=Path)
    args = parser.parse_args()

    root = args.input_dir.expanduser().resolve()
    if not root.is_dir():
        parser.error(f"Input directory does not exist: {root}")

    report = Inventory().build(root)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(
        {"output": str(args.out.resolve()), "files": report["file_count"], "errors": len(report["errors"])},
        ensure_ascii=False,
    ))
    return 1 if report["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
