# 폴더 검사/덤프 반영의 웹/CLI 공통 진입점입니다.
from __future__ import annotations

from pathlib import Path
import traceback
from uuid import uuid4

from app_runtime import BASE_DIR, RESULT_DIR, WORK_DIR, log_event, remove_runtime_path
from output_file_check.folder_apply_ops import (
    apply_dumped_folder,
    copy_folder_to_dump,
    effective_uploaded_root,
    resolve_dump_parent,
)
from output_file_check.folder_mapping import build_folder_mapping, build_folder_policy_from_fields
from output_file_check.folder_serialization import serialize_check_result
from web_uploads import save_check_uploads, save_requirement_uploads


DEFAULT_CHECK_FOLDER = BASE_DIR / "data" / "테스트"


def run_web_check(fields: dict[str, str], file_items: dict[str, list[tuple[str, bytes]]]) -> dict[str, object]:
    # 웹 업로드 요청을 받아 임시 폴더에 저장한 뒤 폴더 매칭만 실행한다.
    """업로드된 폴더와 문서관리표준을 대조하고 웹 화면용 결과를 만든다."""
    WORK_DIR.mkdir(exist_ok=True)
    request_id = uuid4().hex[:8]
    temp_dir = WORK_DIR / f"c-{request_id}"
    temp_dir.mkdir()

    log_event(
        "check.start",
        request_id=request_id,
        fields=list(fields.keys()),
        file_counts={name: len(items) for name, items in file_items.items()},
    )

    try:
        standard_file, folder_dir = save_check_uploads(temp_dir, file_items, fallback_folder=DEFAULT_CHECK_FOLDER)
        payload = run_folder_check_paths(standard_file, folder_dir, fields, request_id=request_id)
        log_check_done(request_id, payload)
        remove_runtime_path(temp_dir)
        return payload
    except Exception as exc:
        log_event(
            "check.error",
            request_id=request_id,
            error=str(exc),
            traceback=traceback.format_exc(),
        )
        remove_runtime_path(temp_dir)
        raise


def run_folder_check_paths(
    standard_file: Path,
    folder_dir: Path,
    fields: dict[str, str] | None = None,
    *,
    request_id: str | None = None,
) -> dict[str, object]:
    # 이미 로컬에 있는 표준 파일/폴더 경로를 기준으로 매칭 결과 JSON을 만든다.
    """웹 업로드 없이 로컬 경로 기준으로 산출물 매칭 결과를 만든다."""
    request_id = request_id or uuid4().hex[:8]
    fields = fields or {}
    folder_policy = build_folder_policy_from_fields(fields)
    mapping = build_folder_mapping(standard_file, folder_dir, fields, folder_policy)
    payload = serialize_check_result(request_id, standard_file, folder_dir, mapping)
    payload["folder_root"] = str(folder_dir)
    return payload


def log_check_done(request_id: str, payload: dict[str, object]) -> None:
    # 긴 검사 요청이 끝났는지 로그만 보고 알 수 있게 요약 완료 로그를 남긴다.
    log_event(
        "check.done",
        request_id=request_id,
        scanned_files=payload.get("scanned_files"),
        output_count=payload.get("output_count"),
        matched_output_count=payload.get("matched_output_count"),
        unmatched_reference_output_count=payload.get("unmatched_reference_output_count"),
    )


def run_web_folder_apply(fields: dict[str, str], file_items: dict[str, list[tuple[str, bytes]]]) -> dict[str, object]:
    # 웹 업로드 폴더를 덤프 위치로 복사한 뒤 복사본에만 문서값 반영을 실행한다.
    """업로드 폴더를 지정 경로에 덤프한 뒤 O 대상 문서만 표준값으로 반영한다."""
    WORK_DIR.mkdir(exist_ok=True)
    RESULT_DIR.mkdir(exist_ok=True)
    request_id = uuid4().hex[:8]
    temp_dir = WORK_DIR / f"a-{request_id}"
    temp_dir.mkdir()

    log_event(
        "folder_apply.start",
        request_id=request_id,
        fields=list(fields.keys()),
        file_counts={name: len(items) for name, items in file_items.items()},
    )

    try:
        standard_file, uploaded_folder = save_check_uploads(temp_dir, file_items, fallback_folder=DEFAULT_CHECK_FOLDER)
        requirement_files = save_requirement_uploads(temp_dir, file_items)
        folder_policy = build_folder_policy_from_fields(fields)
        dump_parent = resolve_dump_parent(fields.get("dump_path", ""))
        source_root = effective_uploaded_root(uploaded_folder)
        dump_root = copy_folder_to_dump(source_root, dump_parent)
        payload = apply_dumped_folder(
            standard_file,
            dump_root,
            fields,
            folder_policy,
            temp_dir,
            request_id,
            log_prefix="folder_apply",
            requirement_files=requirement_files,
        )
        remove_runtime_path(temp_dir)
        return payload
    except Exception as exc:
        log_event(
            "folder_apply.error",
            request_id=request_id,
            error=str(exc),
            traceback=traceback.format_exc(),
        )
        remove_runtime_path(temp_dir)
        raise


def run_folder_apply_paths(
    standard_file: Path,
    source_folder: Path,
    dump_parent: Path | None = None,
    fields: dict[str, str] | None = None,
    *,
    request_id: str | None = None,
) -> dict[str, object]:
    # CLI에서 받은 로컬 폴더를 덤프 복사하고 복사본에 문서값 반영을 실행한다.
    """로컬 폴더를 덤프한 뒤 웹과 같은 방식으로 대상 문서를 일괄 반영한다."""
    WORK_DIR.mkdir(exist_ok=True)
    RESULT_DIR.mkdir(exist_ok=True)
    request_id = request_id or uuid4().hex[:8]
    temp_dir = WORK_DIR / f"a-{request_id}"
    temp_dir.mkdir()
    fields = fields or {}

    log_event(
        "folder_apply_cli.start",
        request_id=request_id,
        standard_file=str(standard_file),
        source_folder=str(source_folder),
        dump_parent=str(dump_parent) if dump_parent else "",
    )

    try:
        if not standard_file.exists():
            raise ValueError(f"문서관리표준 PDF를 찾지 못했습니다: {standard_file}")
        if not source_folder.exists() or not source_folder.is_dir():
            raise ValueError(f"검사 폴더를 찾지 못했습니다: {source_folder}")

        folder_policy = build_folder_policy_from_fields(fields)
        target_parent = dump_parent or resolve_dump_parent("")
        target_parent.mkdir(parents=True, exist_ok=True)
        dump_root = copy_folder_to_dump(source_folder, target_parent)
        payload = apply_dumped_folder(
            standard_file,
            dump_root,
            fields,
            folder_policy,
            temp_dir,
            request_id,
            log_prefix="folder_apply_cli",
        )
        remove_runtime_path(temp_dir)
        return payload
    except Exception as exc:
        log_event(
            "folder_apply_cli.error",
            request_id=request_id,
            error=str(exc),
            traceback=traceback.format_exc(),
        )
        remove_runtime_path(temp_dir)
        raise
