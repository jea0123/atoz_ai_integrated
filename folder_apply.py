# 폴더 산출물 검사/덤프 반영을 CLI에서 실행하는 얇은 진입점입니다.
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from output_file_check.folder_workflow import DEFAULT_CHECK_FOLDER, run_folder_apply_paths, run_folder_check_paths
from output_file_check.matcher import DEFAULT_MATCH_THRESHOLD


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="웹의 산출물 매핑 확인/덤프 후 반영 기능을 CLI에서 실행합니다."
    )
    parser.add_argument("--folder", type=Path, default=DEFAULT_CHECK_FOLDER, help="검사하거나 덤프할 원본 폴더. 생략 시 data\\테스트 사용")
    parser.add_argument("--standard", type=Path, help="문서관리표준 PDF. 생략하면 폴더 안에서 자동 검색")
    parser.add_argument("--dump", type=Path, help="반영 결과 폴더를 만들 상위 경로")
    parser.add_argument("--check-only", action="store_true", help="파일 수정 없이 웹의 확인하기 결과만 출력")
    parser.add_argument("--json", action="store_true", help="전체 결과를 JSON으로 출력")
    parser.add_argument("--threshold", type=float, default=DEFAULT_MATCH_THRESHOLD, help="문서명 매칭 점수 기준")
    parser.add_argument("--project-title", help="표준 PDF에서 읽은 사업명 대신 사용할 사업명")
    parser.add_argument("--ignore-folders", help="매칭에서 제외할 폴더명. 쉼표로 구분")
    parser.add_argument("--transparent-folders", help="경로 비교 때 없는 것처럼 볼 폴더명. 쉼표로 구분")
    parser.add_argument("--map-only-under", help="이 경로 아래만 매핑 대상으로 볼 경로. 쉼표로 구분")
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
            raise RuntimeError("덤프 후 반영을 하려면 --dump 경로를 지정해 주세요.")
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
        # 웹에는 없는 값이다. CLI에서 표준 PDF 사업명 파싱이 틀렸을 때만 덮어쓴다.
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
            raise RuntimeError(f"문서관리표준 PDF를 찾지 못했습니다: {standard_file}")
        return standard_file

    candidates = [path for path in folder.rglob("*문서관리표준*.pdf") if path.is_file()]
    if not candidates:
        raise RuntimeError("--standard를 지정하거나 검사 폴더 안에 문서관리표준 PDF를 넣어 주세요.")

    return max(candidates, key=lambda path: path.stat().st_mtime)


def print_report(payload: dict[str, Any], standard_file: Path, folder: Path, *, check_only: bool) -> None:
    print(f"문서관리표준: {standard_file}")
    print(f"원본 폴더: {folder}")
    if payload.get("dump_root"):
        print(f"덤프 폴더: {payload['dump_root']}")
    if payload.get("standard_project_title"):
        print(f"표준 사업명: {payload['standard_project_title']}")
    print()

    print(f"표준 산출물: {payload.get('output_count', 0)}개")
    print(f"검사 파일: {payload.get('scanned_files', 0)}개")
    print(f"매칭 산출물: {payload.get('matched_output_count', 0)}개")
    print(f"매칭 파일: {payload.get('matched_file_count', 0)}개")
    print(f"미매칭 산출물: {payload.get('missing_output_count', 0)}개")

    if check_only:
        print("\n검사만 실행했습니다. 파일은 수정하지 않았습니다.")
        print_match_preview(payload)
        return

    print(f"반영 대상 파일: {payload.get('apply_target_file_count', 0)}개")
    print(f"반영 성공: {payload.get('updated_file_count', 0)}개")
    print(f"반영 오류: {payload.get('failed_file_count', 0)}개")
    print_apply_errors(payload)


def print_match_preview(payload: dict[str, Any], limit: int = 20) -> None:
    matches = payload.get("matches") or []
    if not matches:
        return

    print("\n매칭 결과")
    for item in matches[:limit]:
        candidates = item.get("candidates") or []
        print(f"- {item.get('output_name', '')} ({item.get('output_id', '')}): {len(candidates)}개")
    if len(matches) > limit:
        print(f"... {len(matches) - limit}개 더 있음")


def print_apply_errors(payload: dict[str, Any]) -> None:
    errors = [
        item for item in payload.get("apply_items", [])
        if isinstance(item, dict) and item.get("status") == "error"
    ]
    if not errors:
        return

    print("\n오류")
    for item in errors:
        print(f"- {item.get('output_name', '')} ({item.get('output_id', '')})")
        print(f"  위치: {item.get('old_path', '')}")
        print(f"  내용: {item.get('error', '')}")
        if item.get("backup_path"):
            print(f"  백업: {item.get('backup_path')}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
