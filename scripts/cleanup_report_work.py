#!/usr/bin/env python3
"""Safely remove task-created report intermediates after final QA."""

from __future__ import annotations

import argparse
from pathlib import Path

from report_cleanup.plan import CleanupPlan


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--final", required=True, type=Path, help="Validated final DOCX")
    parser.add_argument("--work-dir", required=True, type=Path, help="Dedicated task work directory")
    parser.add_argument("--draft", action="append", default=[], type=Path, help="Task-created draft file")
    parser.add_argument("--apply", action="store_true", help="Delete instead of previewing")
    args = parser.parse_args()

    final = args.final.expanduser().resolve()
    work = args.work_dir.expanduser().resolve()
    drafts = [path.expanduser().resolve() for path in args.draft]
    plan = CleanupPlan(final, work, drafts)
    if args.apply:
        plan.apply()
    else:
        plan.preview()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
