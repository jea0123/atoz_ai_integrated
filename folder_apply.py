from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from output_file_check.folder_workflow import DEFAULT_CHECK_FOLDER, run_folder_apply_paths, run_folder_check_paths
from output_file_check.matcher import DEFAULT_MATCH_THRESHOLD


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run output-folder matching or apply from the CLI.")
    parser.add_argument("--folder", type=Path, default=DEFAULT_CHECK_FOLDER, help="Source folder to check/apply.")
    parser.add_argument("--standard", type=Path, help="Document standard PDF. If omitted, search under --folder.")
    parser.add_argument("--dump", type=Path, help="Parent folder for apply result dumps.")
    parser.add_argument("--check-only", action="store_true", help="Only print matching results.")
    parser.add_argument("--json", action="store_true", help="Print full JSON payload.")
    parser.add_argument("--threshold", type=float, default=DEFAULT_MATCH_THRESHOLD, help="Match score threshold.")
    parser.add_argument("--project-title", help="Override project title read from the standard.")
    parser.add_argument("--ignore-folders", help="Comma-separated folder names to ignore.")
    parser.add_argument("--transparent-folders", help="Comma-separated folder names ignored in path comparison.")
    parser.add_argument("--map-only-under", help="Comma-separated path prefixes to scan.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    folder = args.folder.expanduser().resolve()
    standard_file = resolve_standard_file(args.standard, folder)
    fields = build_fields(args)

    if args.check_only:
        payload = run_folder_check_paths(standard_file, folder, fields)
    else:
        if args.dump is None:
            raise RuntimeError("--dump is required unless --check-only is used.")
        payload = run_folder_apply_paths(
            standard_file,
            folder,
            args.dump.expanduser().resolve(),
            fields,
        )

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print_report(payload, standard_file, folder, check_only=args.check_only)
    return 0


def build_fields(args: argparse.Namespace) -> dict[str, str]:
    fields = {"threshold": str(args.threshold)}
    optional_values = {
        "project_title": args.project_title,
        "ignore_folder_names": args.ignore_folders,
        "transparent_folder_names": args.transparent_folders,
        "map_only_under": args.map_only_under,
    }
    fields.update({key: value for key, value in optional_values.items() if value})
    return fields


def resolve_standard_file(supplied: Path | None, folder: Path) -> Path:
    if supplied:
        standard_file = supplied.expanduser().resolve()
        if not standard_file.exists():
            raise RuntimeError(f"Standard file not found: {standard_file}")
        return standard_file

    candidates = [path for path in folder.rglob("*문서관리표준*.pdf") if path.is_file()]
    if not candidates:
        raise RuntimeError("Pass --standard or place a 문서관리표준 PDF under the source folder.")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def print_report(payload: dict[str, Any], standard_file: Path, folder: Path, *, check_only: bool) -> None:
    print(f"Standard: {standard_file}")
    print(f"Source folder: {folder}")
    if payload.get("dump_root"):
        print(f"Dump folder: {payload['dump_root']}")
    if payload.get("standard_project_title"):
        print(f"Project title: {payload['standard_project_title']}")
    print()

    print(f"Reference outputs: {payload.get('output_count', 0)}")
    print(f"Scanned files: {payload.get('scanned_files', 0)}")
    print(f"Matched outputs: {payload.get('matched_output_count', 0)}")
    print(f"Matched files: {payload.get('matched_file_count', 0)}")
    print(f"Unmatched outputs: {payload.get('unmatched_reference_output_count', 0)}")

    if check_only:
        print("\nCheck-only mode. Files were not modified.")
        print_match_preview(payload)
        return

    print(f"Apply target files: {payload.get('apply_target_file_count', 0)}")
    print(f"Updated: {payload.get('updated_file_count', 0)}")
    print(f"Failed: {payload.get('failed_file_count', 0)}")
    print_apply_errors(payload)


def print_match_preview(payload: dict[str, Any], limit: int = 20) -> None:
    matches = payload.get("matches") or []
    if not matches:
        return

    print("\nMatches")
    for item in matches[:limit]:
        candidates = item.get("candidates") or []
        print(f"- {item.get('output_name', '')} ({item.get('output_id', '')}): {len(candidates)}")
    if len(matches) > limit:
        print(f"... {len(matches) - limit} more")


def print_apply_errors(payload: dict[str, Any]) -> None:
    errors = [
        item for item in payload.get("apply_items", [])
        if isinstance(item, dict) and item.get("status") == "error"
    ]
    if not errors:
        return

    print("\nErrors")
    for item in errors:
        print(f"- {item.get('output_name', '')} ({item.get('output_id', '')})")
        print(f"  Path: {item.get('old_path', '')}")
        print(f"  Error: {item.get('error', '')}")
        if item.get("backup_path"):
            print(f"  Backup: {item.get('backup_path')}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
