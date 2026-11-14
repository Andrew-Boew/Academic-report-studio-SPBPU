#!/usr/bin/env python3
"""Build a Russian academic report DOCX from a constrained JSON specification."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from report_build.builder import build
from report_build.errors import SpecError
from report_build.spec.metadata import MetadataMerger


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--spec", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--profile", type=Path, help="Optional reusable .report-profile.json")
    parser.add_argument("--subjects", type=Path, help="Optional discipline reviewers .report-subjects.json")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()

    spec_path = args.spec.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if output.suffix.lower() != ".docx":
        parser.error("The final output must use the .docx extension")
    if output.exists() and not args.force:
        parser.error(f"Output already exists; pass --force to replace it: {output}")
    try:
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        built_in_assets = Path(__file__).resolve().parent.parent / "assets"
        profile_path = args.profile
        if profile_path is None and isinstance(spec, dict) and spec.get("profile_path"):
            profile_path = Path(str(spec["profile_path"]))
        if profile_path is None:
            candidate = built_in_assets / "report-profile.default.json"
            if candidate.is_file():
                profile_path = candidate
        subjects_path = args.subjects
        if subjects_path is None and isinstance(spec, dict) and spec.get("subject_profile_path"):
            subjects_path = Path(str(spec["subject_profile_path"]))
        current_metadata = spec.get("metadata") if isinstance(spec, dict) else None
        cover_layout = str((current_metadata or {}).get("cover_layout", "")).strip().casefold()
        reviewer_is_missing = not isinstance(current_metadata, dict) or any(
            not str(current_metadata.get(key, "")).strip()
            for key in ("reviewer_position", "reviewer_name")
        )
        if subjects_path is None and reviewer_is_missing and cover_layout not in {"vkr_2026", "vkr_tr23", "nir"}:
            candidate = built_in_assets / "report-subjects.default.json"
            if candidate.is_file():
                subjects_path = candidate
        spec = MetadataMerger().merge_metadata_sources(spec, profile_path, subjects_path)
        build(spec, output)
    except (OSError, json.JSONDecodeError, SpecError) as exc:
        parser.error(str(exc))
    print(json.dumps({"output": str(output), "format": "docx"}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
