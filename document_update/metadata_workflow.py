from __future__ import annotations

import os
from pathlib import Path
import re
import shutil

from app_runtime import RESULT_DIR
from output_file_check.standard_reader import extract_standard_text
from web_uploads import safe_relative_upload_path, safe_upload_filename

from .metadata_update import (
    IGNORED_FOLDER_NAMES,
    MetadataTarget,
    build_metadata_targets,
    read_wbs_metadata,
    update_metadata_in_document,
)


WBS_SUFFIXES = {".xlsx", ".xlsm"}
STANDARD_SUFFIXES = {".pdf", ".hwp", ".hwpx"}
METADATA_UPLOAD_SUFFIXES = {".hwp", ".hwpx", ".xlsx", ".xlsm", ".xltx", ".xltm"}


def effective_uploaded_root(folder_dir: Path) -> Path:
    children = list(folder_dir.iterdir())
    directories = [path for path in children if path.is_dir()]
    files = [path for path in children if path.is_file()]
    if len(directories) == 1 and not files:
        return directories[0]
    return folder_dir


def save_metadata_inputs(
    temp_dir: Path,
    file_items: dict[str, list[tuple[str, bytes]]],
) -> tuple[Path, Path, Path, bool]:
    wbs_path, standard_path = save_metadata_required_files(temp_dir, file_items)
    uploaded_folder = save_metadata_folder(temp_dir, file_items)
    if uploaded_folder is not None:
        return wbs_path, standard_path, effective_uploaded_root(uploaded_folder), True
    raise ValueError("산출물 폴더를 선택하세요.")


def save_metadata_required_files(
    temp_dir: Path,
    file_items: dict[str, list[tuple[str, bytes]]],
) -> tuple[Path, Path]:
    wbs_path = save_wbs_file(temp_dir, file_items)
    if wbs_path is None:
        raise ValueError("WBS 파일을 선택하세요. 지원 확장자: .xlsx, .xlsm")

    standard_path = save_standard_file(temp_dir, file_items)
    if standard_path is None:
        raise ValueError("문서관리표준 파일을 선택하세요. 지원 확장자: .pdf, .hwp, .hwpx")

    return wbs_path, standard_path


def resolve_existing_dump_root(raw_path: str) -> Path:
    value = raw_path.strip().strip('"')
    if not value:
        raise ValueError("산출물 매핑 결과 폴더 경로를 입력하세요.")
    path = Path(value).expanduser()
    if not path.is_dir():
        raise ValueError(f"산출물 매핑 결과 폴더를 찾지 못했습니다: {path}")
    return path


def save_wbs_file(temp_dir: Path, file_items: dict[str, list[tuple[str, bytes]]]) -> Path | None:
    items = file_items.get("wbs_file") or []
    if not items or not items[0][1]:
        return None
    filename, payload = items[0]
    suffix = Path(filename).suffix.lower()
    if suffix not in WBS_SUFFIXES:
        raise ValueError("WBS 파일은 .xlsx 또는 .xlsm만 사용할 수 있습니다.")
    path = temp_dir / safe_upload_filename(filename, "wbs_file", suffix)
    path.write_bytes(payload)
    return path


def save_standard_file(temp_dir: Path, file_items: dict[str, list[tuple[str, bytes]]]) -> Path | None:
    items = file_items.get("standard_file") or []
    if not items or not items[0][1]:
        return None
    filename, payload = items[0]
    suffix = Path(filename).suffix.lower()
    if suffix not in STANDARD_SUFFIXES:
        raise ValueError("문서관리표준은 .pdf, .hwp, .hwpx만 사용할 수 있습니다.")
    path = temp_dir / safe_upload_filename(filename, "standard_file", suffix)
    path.write_bytes(payload)
    return path


def read_standard_cover_author(standard_path: Path) -> str:
    text = extract_standard_text(standard_path)[:3000]
    match = re.search(r"작성\s*자\s+([^\s<>\|]+)", text)
    if match:
        return match.group(1).strip()
    raise RuntimeError(f"문서관리표준 표지에서 작성자를 찾지 못했습니다: {standard_path}")


def save_metadata_folder(temp_dir: Path, file_items: dict[str, list[tuple[str, bytes]]]) -> Path | None:
    items = file_items.get("document_files") or []
    if not items:
        return None

    folder_dir = temp_dir / "uploaded-documents"
    folder_dir.mkdir()
    saved_count = 0
    for index, (filename, payload) in enumerate(items, start=1):
        if not payload:
            continue
        suffix = Path(filename).suffix.lower()
        if suffix not in METADATA_UPLOAD_SUFFIXES:
            continue
        relative_path = safe_relative_upload_path(filename, f"document-{index}{suffix}")
        target_path = folder_dir / relative_path
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(payload)
        saved_count += 1

    if saved_count == 0:
        raise ValueError("선택한 폴더에서 수정 가능한 문서를 찾지 못했습니다.")
    return folder_dir


def serialize_target(target: MetadataTarget) -> dict[str, object]:
    return {
        "path": str(target.path),
        "relative_path": target.relative_path,
        "status": target.status,
        "message": target.message,
        "output_name": target.output_name,
        "author": target.author,
        "revision_date": target.revision_date,
        "current": {
            "author": target.current.author,
            "revision_date": target.current.revision_date,
            "revision_author": target.current.revision_author,
            "revision_history_date": target.current.revision_history_date,
        },
        "candidate_count": len(target.candidates),
        "candidates": [
            {
                "output_name": item.output_name,
                "author": item.author,
                "revision_date": item.revision_date,
                "wbs": item.wbs,
                "task": item.task,
                "row": item.row,
                "requirement_id": item.requirement_id,
            }
            for item in target.candidates[:12]
        ],
    }


def run_metadata_preview(wbs_path: Path, standard_path: Path, folder_root: Path, request_id: str) -> dict[str, object]:
    records = read_wbs_metadata(wbs_path)
    approval_author = read_standard_cover_author(standard_path)
    targets = build_metadata_targets(folder_root, records)
    matched = [target for target in targets if target.status == "matched"]
    ambiguous = [target for target in targets if target.status == "ambiguous"]
    unmatched = [target for target in targets if target.status == "unmatched"]
    return {
        "ok": True,
        "request_id": request_id,
        "wbs_path": str(wbs_path),
        "standard_path": str(standard_path),
        "approval_author": approval_author,
        "folder_root": str(folder_root),
        "wbs_record_count": len(records),
        "document_count": len(targets),
        "matched_count": len(matched),
        "ambiguous_count": len(ambiguous),
        "unmatched_count": len(unmatched),
        "targets": [serialize_target(target) for target in targets],
    }


def run_metadata_apply(
    wbs_path: Path,
    standard_path: Path,
    source_root: Path,
    request_id: str,
    excluded_paths: set[str],
) -> dict[str, object]:
    dump_parent = RESULT_DIR / "metadata-dumps"
    dump_parent.mkdir(parents=True, exist_ok=True)
    dump_root = next_dump_path(dump_parent, source_root.name)
    shutil.copytree(windows_long_path(source_root), windows_long_path(dump_root), ignore=ignore_metadata_dirs)

    payload = apply_metadata_to_existing_dump(
        wbs_path,
        standard_path,
        dump_root,
        request_id,
        excluded_paths,
        temp_parent=dump_parent,
    )
    payload["source_root"] = str(source_root)
    payload["metadata_copy_created"] = True
    return payload


def apply_metadata_to_existing_dump(
    wbs_path: Path,
    standard_path: Path,
    dump_root: Path,
    request_id: str,
    excluded_paths: set[str],
    *,
    temp_parent: Path | None = None,
) -> dict[str, object]:
    if not dump_root.is_dir():
        raise ValueError(f"후처리 대상 폴더를 찾지 못했습니다: {dump_root}")

    temp_parent = temp_parent or (RESULT_DIR / "metadata-dumps")
    temp_parent.mkdir(parents=True, exist_ok=True)

    records = read_wbs_metadata(wbs_path)
    approval_author = read_standard_cover_author(standard_path)
    targets = build_metadata_targets(dump_root, records)
    apply_targets = [
        target
        for target in targets
        if target.status == "matched" and normalize_relative_path(target.relative_path) not in excluded_paths
    ]

    temp_dir = temp_parent / f".metadata-temp-{request_id}"
    temp_dir.mkdir(exist_ok=True)
    items = []
    try:
        for target in apply_targets:
            result = update_metadata_in_document(
                target.path,
                target.author,
                target.revision_date,
                approval_author,
                temp_dir,
            )
            items.append(
                {
                    "status": result.status,
                    "old_path": str(result.old_path),
                    "new_path": str(result.new_path or ""),
                    "backup_path": str(result.backup_path or ""),
                    "converted_to_hwpx": result.converted_to_hwpx,
                    "cover_update_count": result.cover_update_count,
                    "revision_history_update_count": result.revision_history_update_count,
                    "error": result.error,
                    "relative_path": target.relative_path,
                    "output_name": target.output_name,
                    "author": target.author,
                    "revision_date": target.revision_date,
                    "approval_author": approval_author,
                }
            )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    failed = [item for item in items if item["status"] == "error"]
    return {
        "ok": not failed and bool(items),
        "request_id": request_id,
        "standard_path": str(standard_path),
        "approval_author": approval_author,
        "dump_root": str(dump_root),
        "metadata_copy_created": False,
        "updated_file_count": sum(1 for item in items if item["status"] == "updated"),
        "failed_file_count": len(failed),
        "skipped_file_count": len(targets) - len(apply_targets),
        "apply_target_file_count": len(apply_targets),
        "apply_items": items,
        "targets": [serialize_target(target) for target in targets],
    }


def normalize_relative_path(value: str) -> str:
    return value.replace("\\", "/").strip().casefold()


def split_excluded_paths(raw: str) -> set[str]:
    return {normalize_relative_path(line) for line in raw.splitlines() if line.strip()}


def ignore_metadata_dirs(_directory: str, names: list[str]) -> set[str]:
    return {name for name in names if name in IGNORED_FOLDER_NAMES}


def windows_long_path(path: Path) -> str:
    if os.name != "nt":
        return str(path)

    text = str(path)
    if text.startswith("\\\\?\\"):
        return text

    absolute = str(path.resolve(strict=False))
    if absolute.startswith("\\\\?\\"):
        return absolute
    if absolute.startswith("\\\\"):
        return "\\\\?\\UNC\\" + absolute.lstrip("\\")
    return "\\\\?\\" + absolute


def next_dump_path(parent: Path, folder_name: str) -> Path:
    base = f"{folder_name}_metadata"
    pattern = re.compile(rf"^{re.escape(base)}_v0\.(\d+)$", re.IGNORECASE)
    highest_minor = 0
    for child in parent.iterdir():
        if not child.is_dir():
            continue
        match = pattern.fullmatch(child.name)
        if match:
            highest_minor = max(highest_minor, int(match.group(1)))

    index = highest_minor + 1
    candidate = parent / f"{base}_v0.{index}"
    while candidate.exists():
        index += 1
        candidate = parent / f"{base}_v0.{index}"
    return candidate
