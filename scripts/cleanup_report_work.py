#!/usr/bin/env python3
"""Safely remove task-created report intermediates after final QA."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path


def inside(child: Path, parent: Path) -> bool:
    try:
        child.relative_to(parent)
        return True
    except ValueError:
        return False


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
    if not final.is_file() or final.suffix.lower() != ".docx" or final.stat().st_size == 0:
        raise SystemExit("Final DOCX is missing or empty; cleanup refused")
    if work == work.parent or len(work.parts) < 4:
        raise SystemExit(f"Work directory is too broad: {work}")
    if inside(final, work):
        raise SystemExit("Final DOCX must be outside the disposable work directory")
    for draft in drafts:
        if draft == final:
            raise SystemExit("A draft path resolves to the final DOCX")

    targets = [path for path in [work, *drafts] if path.exists()]
    for target in targets:
        print(("DELETE" if args.apply else "WOULD_DELETE") + f"\t{target}")
    if not args.apply:
        return 0
    for target in targets:
        if target.is_dir() and not target.is_symlink():
            shutil.rmtree(target)
        else:
            target.unlink()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
