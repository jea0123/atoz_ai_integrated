# 덤프 폴더 복사, 문서값 반영, 파일명 변경을 실행합니다.
from __future__ import annotations

from datetime import datetime
import os
from pathlib import Path
import re
import shutil
from uuid import uuid4

from app_runtime import (
    RESULT_DIR,
    log_event,
)
from document_update.document_number import write_updated_document
from document_update.hwp_convert import start_allow_all_watcher, stop_allow_all_watcher
from document_update.patterns import OUTPUT_ID_PATTERN
from document_update.runtime_conversion import prepare_target_file
from output_file_check.file_noise import copytree_ignore_noise, remove_noise_files
from output_file_check.folder_mapping import (
    build_folder_mapping,
    normalize_relative_path_for_compare,
    split_excluded_paths_field,
)
from output_file_check.folder_policy import FolderPolicy
from output_file_check.folder_serialization import serialize_check_result
from output_file_check.models import MatchCandidate, OutputMatch, PathTemplate, StandardOutput
from output_file_check.normalization import filesystem_safe_stem, normalize_for_match, output_name_from_id
from output_file_check.requirement_generation import (
    RequirementGenerationResult,
    generate_requirement_documents,
)


TEST_OUTPUT_ID_PATTERN = re.compile(r"\d{4,}(?:\([^)]*\))?")
PRESERVED_TAIL_PATTERN = re.compile(
    r"((?:(?:[_-][vV]\d+(?:\.\d+)*)|(?:[_-]SFR-[A-Za-z0-9-]+))+)$",
    re.IGNORECASE,
)
ATTACHMENT_TAIL_PATTERN = re.compile(
    r"^[\s_-]*(?:[\[\(（［｛]\s*(?:별첨|첨부)\s*\d*[^)\]\}）］｝]*[\)\]\}）］｝]|(?:별첨|첨부)\s*\d+)",
    re.IGNORECASE,
)
REQUIREMENT_TAIL_PATTERN = re.compile(r"(?<![A-Z0-9])SFR-(?:[A-Z0-9]+-)*\d+(?![A-Z0-9])", re.IGNORECASE)
VERSION_TAIL_PATTERN = re.compile(r"([_-])[vV]\d+(?:\.\d+)*")
DEFAULT_FILENAME_VERSION_TAIL = "_v0.1"
HANDOVER_RESULT_KEY = normalize_for_match("인수인계시험결과서")
HANDOVER_CONFIRMATION_KEY = normalize_for_match("별첨1인수인계확인서")


def clean_output_id(raw_text: str) -> str:
    text = raw_text.strip()
    if not text:
        return ""

    match = OUTPUT_ID_PATTERN.search(text)
    if match:
        return match.group(0)

    return text.strip(" -\t\n\r")


def validate_output_id(output_id: str) -> None:
    if not OUTPUT_ID_PATTERN.fullmatch(output_id) and not TEST_OUTPUT_ID_PATTERN.fullmatch(output_id):
        raise RuntimeError(f"산출물 ID 형식이 예상과 다릅니다: {output_id}")


def build_target_filename(output: StandardOutput, old_path: Path) -> str:
    suffix = old_path.suffix
    stem = build_standard_stem(output)
    preserved_tail = extract_preserved_tail(old_path.stem, stem)
    return f"{filesystem_safe_stem(f'{stem}{preserved_tail}')}{suffix}"


def build_standard_stem(output: StandardOutput) -> str:
    id_name = output_name_from_id(output.output_id)
    output_name = output.output_name

    if id_name and normalize_for_match(id_name) == normalize_for_match(output_name):
        return output.output_id

    if id_name:
        return output.output_id

    return f"{output.output_id}-{output_name}" if output.output_id else output_name


def extract_preserved_tail(stem: str, standard_stem: str = "") -> str:
    full_tail = tail_after_standard_stem(stem, standard_stem)
    if (
        full_tail is not None
        and (
            is_attachment_tail(full_tail)
            or (
                is_handover_confirmation_stem(stem, standard_stem)
                and REQUIREMENT_TAIL_PATTERN.search(full_tail)
            )
        )
    ):
        tail = full_tail
    else:
        match = PRESERVED_TAIL_PATTERN.search(stem)
        if not match:
            return DEFAULT_FILENAME_VERSION_TAIL
        tail = match.group(1)

    tail = VERSION_TAIL_PATTERN.sub(r"\1v0.1", tail)
    if not VERSION_TAIL_PATTERN.search(tail):
        tail = f"{tail}{DEFAULT_FILENAME_VERSION_TAIL}"
    if tail and tail[0] not in {"_", "-", "["}:
        tail = f"_{tail}"
    return tail


def is_attachment_tail(tail: str) -> bool:
    return bool(ATTACHMENT_TAIL_PATTERN.search(tail))


def is_handover_confirmation_stem(stem: str, standard_stem: str) -> bool:
    stem_key = normalize_for_match(stem)
    standard_key = normalize_for_match(standard_stem)
    return HANDOVER_RESULT_KEY in standard_key and HANDOVER_CONFIRMATION_KEY in stem_key


def tail_after_standard_stem(stem: str, standard_stem: str) -> str | None:
    if not standard_stem:
        return None

    prefixes = tuple(dict.fromkeys([standard_stem, filesystem_safe_stem(standard_stem)]))
    for prefix in prefixes:
        if not prefix or len(stem) < len(prefix):
            continue
        if stem[:len(prefix)].casefold() != prefix.casefold():
            continue
        tail = stem[len(prefix):]
        if not tail or tail[0] in {"_", "-", "[", " "}:
            return tail
    return None


def copy_folder_to_dump(source_root: Path, dump_parent: Path) -> Path:
    # 원본 폴더를 건드리지 않기 위해 덤프 위치에 복사본을 만든다.
    dump_root = next_versioned_dump_path(dump_parent, source_root.name)
    shutil.copytree(
        windows_long_path(source_root),
        windows_long_path(dump_root),
        ignore=copytree_ignore_noise,
    )
    remove_noise_files(dump_root)
    return dump_root


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


def apply_dumped_folder(
    standard_file: Path,
    dump_root: Path,
    fields: dict[str, str],
    folder_policy: FolderPolicy,
    temp_dir: Path,
    request_id: str,
    *,
    log_prefix: str,
    requirement_files: list[Path] | None = None,
) -> dict[str, object]:
    # 덤프된 폴더 기준으로 매칭, 제외 후보 반영, 문서 수정, 파일명 변경을 실행한다.
    mapping = build_folder_mapping(standard_file, dump_root, fields, folder_policy)

    all_selected_candidates = select_all_unique_candidates(mapping.matches)
    selected_candidates = filter_excluded_candidates(
        all_selected_candidates,
        dump_root,
        split_excluded_paths_field(fields.get("excluded_candidate_paths")),
    )
    excluded_file_count = len(all_selected_candidates) - len(selected_candidates)
    allow_stop_event, allow_thread = start_allow_all_watcher()
    try:
        apply_items = [
            apply_batch_candidate(
                candidate,
                mapping.standard_project_title,
                temp_dir,
                rename_files=True,
            )
            for candidate in selected_candidates
        ]
    finally:
        stop_allow_all_watcher(allow_stop_event, allow_thread)

    requirement_result = run_requirement_generation_safely(
        dump_root,
        mapping.outputs,
        mapping.path_templates,
        requirement_files or [],
        mapping.standard_project_title,
        temp_dir,
        fields,
        apply_items=apply_items,
        request_id=request_id,
        log_prefix=log_prefix,
    )
    removed_noise_file_count = remove_noise_files(dump_root)

    try:
        write_apply_readme(
            dump_root,
            standard_file,
            mapping.standard_project_title,
            apply_items,
            excluded_file_count,
            requirement_result=requirement_result,
        )
    except Exception as exc:
        log_event(
            f"{log_prefix}.readme_error",
            request_id=request_id,
            dump_root=str(dump_root),
            error=str(exc),
        )

    payload = serialize_check_result(
        request_id,
        standard_file,
        dump_root,
        mapping,
    )
    payload.update(
        {
            "dump_root": str(dump_root),
            "updated_file_count": sum(1 for item in apply_items if item["status"] == "updated"),
            "failed_file_count": sum(1 for item in apply_items if item["status"] == "error"),
            "apply_target_file_count": len(selected_candidates),
            "skipped_file_count": excluded_file_count,
            "filename_unchanged_count": sum(1 for item in apply_items if is_filename_unchanged_item(item)),
            "apply_items": apply_items,
            "removed_noise_file_count": removed_noise_file_count,
            **serialize_requirement_generation_result(requirement_result),
        }
    )
    log_event(
        f"{log_prefix}.done",
        request_id=request_id,
        dump_root=str(dump_root),
        selected=len(selected_candidates),
        excluded=excluded_file_count,
        updated=payload["updated_file_count"],
        failed=payload["failed_file_count"],
        removed_noise_file_count=removed_noise_file_count,
        failed_items=[
            {
                "output_name": item.get("output_name", ""),
                "old_path": item.get("old_path", ""),
                "error": item.get("error", ""),
            }
            for item in apply_items
            if item.get("status") == "error"
        ]
    )
    return payload


def run_requirement_generation_safely(
    dump_root: Path,
    outputs: list[StandardOutput],
    path_templates: list[PathTemplate],
    requirement_files: list[Path],
    standard_project_title: str,
    temp_dir: Path,
    fields: dict[str, str],
    *,
    apply_items: list[dict[str, object]] | None = None,
    request_id: str,
    log_prefix: str,
) -> RequirementGenerationResult:
    try:
        return generate_requirement_documents(
            dump_root,
            outputs,
            path_templates,
            requirement_files,
            standard_project_title,
            temp_dir,
            fields,
            apply_items=apply_items,
        )
    except Exception as exc:
        log_event(
            f"{log_prefix}.requirement_generation_error",
            request_id=request_id,
            dump_root=str(dump_root),
            error=str(exc),
        )
        return RequirementGenerationResult(
            enabled=bool(requirement_files),
            target_names=(),
            target_count=0,
            created_items=[],
            skipped_items=[],
            error_items=[
                {
                    "status": "error",
                    "reason": "요구사항별 자동 생성 처리 중 오류가 발생했습니다.",
                    "error": str(exc),
                }
            ],
        )


def serialize_requirement_generation_result(result: RequirementGenerationResult) -> dict[str, object]:
    created_ok = [item for item in result.created_items if item.get("status") != "error"]
    created_errors = [item for item in result.created_items if item.get("status") == "error"]
    warnings = [item for item in result.created_items if item.get("status") == "created_with_warning"]
    removed_items = result.removed_items or []
    removed_ok = [item for item in removed_items if item.get("status") == "removed"]
    removed_errors = [item for item in removed_items if item.get("status") == "error"]
    folder_items = result.folder_items or []
    folder_ok = [item for item in folder_items if item.get("status") != "error"]
    folder_created = [item for item in folder_items if item.get("status") == "created"]
    folder_errors = [item for item in folder_items if item.get("status") == "error"]
    return {
        "requirement_generation_enabled": result.enabled,
        "requirement_generation_target_count": result.target_count,
        "requirement_generated_file_count": len(created_ok),
        "requirement_generated_folder_count": len(folder_ok),
        "requirement_generation_created_folder_count": len(folder_created),
        "requirement_generation_removed_file_count": len(removed_ok),
        "requirement_generation_warning_count": len(warnings),
        "requirement_generation_skipped_count": len(result.skipped_items),
        "requirement_generation_error_count": len(created_errors) + len(result.error_items) + len(removed_errors) + len(folder_errors),
        "requirement_generation_readme_path": str(result.readme_path) if result.readme_path else "",
        "requirement_generation_items": result.created_items,
        "requirement_generation_folder_items": folder_items,
        "requirement_generation_removed_items": removed_items,
        "requirement_generation_skipped_items": result.skipped_items,
        "requirement_generation_error_items": [*created_errors, *result.error_items, *removed_errors, *folder_errors],
    }


def resolve_dump_parent(raw_path: str) -> Path:
    # 사용자가 입력한 덤프 경로가 없으면 web_runtime 아래 공식 결과 위치를 만든다.
    value = raw_path.strip()
    if value:
        path = Path(value).expanduser()
    else:
        path = RESULT_DIR / "folder-dumps"
    path.mkdir(parents=True, exist_ok=True)
    return path


def effective_uploaded_root(folder_dir: Path) -> Path:
    # 브라우저 업로드가 루트 폴더를 한 겹 더 만든 경우 실제 원본 루트를 찾는다.
    children = list(folder_dir.iterdir())
    directories = [path for path in children if path.is_dir()]
    files = [path for path in children if path.is_file()]
    if len(directories) == 1 and not files:
        return directories[0]
    return folder_dir


def next_versioned_dump_path(parent: Path, folder_name: str) -> Path:
    # 반복 실행 결과가 한눈에 보이도록 폴더명 뒤에 v0.1, v0.2 순번을 붙인다.
    pattern = re.compile(rf"^{re.escape(folder_name)}_v0\.(\d+)$", re.IGNORECASE)
    highest_minor = 0
    for child in parent.iterdir():
        if not child.is_dir():
            continue
        match = pattern.fullmatch(child.name)
        if match:
            highest_minor = max(highest_minor, int(match.group(1)))

    index = highest_minor + 1
    candidate = parent / f"{folder_name}_v0.{index}"
    while candidate.exists():
        index += 1
        candidate = parent / f"{folder_name}_v0.{index}"
    return candidate


def unique_file_path(path: Path) -> Path:
    # 백업/출력 파일명이 겹치면 뒤에 번호를 붙여 충돌을 피한다.
    if not path.exists():
        return path

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = path.with_name(f"{path.stem}_{timestamp}{path.suffix}")
    index = 1
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}_{timestamp}_{index}{path.suffix}")
        index += 1
    return candidate


def select_all_unique_candidates(matches: list[OutputMatch]) -> list[MatchCandidate]:
    # 한 파일이 여러 산출물 후보로 잡힌 경우 가장 높은 점수 후보 하나만 남긴다.
    candidates_by_file: dict[Path, MatchCandidate] = {}
    for match in matches:
        for candidate in match.candidates:
            existing = candidates_by_file.get(candidate.file.path)
            if existing is None or candidate.score > existing.score:
                candidates_by_file[candidate.file.path] = candidate
    return sorted(candidates_by_file.values(), key=lambda candidate: str(candidate.file.path).casefold())


def filter_excluded_candidates(
    candidates: list[MatchCandidate],
    dump_root: Path,
    excluded_relative_paths: set[str],
) -> list[MatchCandidate]:
    # 화면에서 '반영 제외' 체크한 후보 파일을 실제 반영 목록에서 뺀다.
    if not excluded_relative_paths:
        return candidates

    filtered: list[MatchCandidate] = []
    for candidate in candidates:
        try:
            relative_path = str(candidate.file.path.relative_to(dump_root))
        except ValueError:
            relative_path = str(candidate.file.path)
        if normalize_relative_path_for_compare(relative_path) in excluded_relative_paths:
            continue
        filtered.append(candidate)
    return filtered


def apply_batch_candidate(
    candidate: MatchCandidate,
    standard_project_title: str,
    temp_dir: Path,
    *,
    rename_files: bool,
) -> dict[str, object]:
    # 후보 파일 하나에 산출물 ID/제목/프로젝트명을 반영하고 필요하면 파일명도 바꾼다.
    original_path = candidate.file.path
    identity = candidate.file.identity
    backup_path = None

    try:
        target_file, _ = prepare_target_file(original_path, temp_dir)
        update_dir = temp_dir / "batch-updated"
        update_dir.mkdir(parents=True, exist_ok=True)
        output_suffix = target_file.suffix or original_path.suffix
        temp_output = unique_file_path(update_dir / f"{uuid4().hex}{output_suffix}")
        old_title = identity.document_title if identity and identity.document_title else None
        old_project_title = identity.project_title if identity and identity.project_title else None
        effective_output = candidate.output
        output_id = clean_output_id(effective_output.output_id)
        validate_output_id(output_id)

        (
            old_document_number,
            _document_backup_path,
            title_replace_count,
            project_title_replace_count,
            document_number_replace_count,
            output_file,
        ) = write_updated_document(
            target_file,
            new_document_number=output_id,
            old_title=old_title,
            new_title=effective_output.output_name if old_title else None,
            old_project_title=old_project_title,
            new_project_title=standard_project_title,
            output_path=temp_output,
        )

        final_path = original_path if original_path.suffix.lower() == output_suffix.lower() else original_path.with_suffix(output_suffix)
        expected_filename = build_target_filename(effective_output, final_path)
        if rename_files:
            final_path = final_path.with_name(expected_filename)
        if final_path.exists() and final_path.resolve() != original_path.resolve():
            final_path = unique_file_path(final_path)

        final_path.parent.mkdir(parents=True, exist_ok=True)
        output_file.replace(final_path)
        if original_path.exists() and original_path.resolve() != final_path.resolve():
            original_path.unlink()

        cover_status = build_cover_status(
            old_document_number=old_document_number,
            new_document_number=output_id,
            old_title=old_title,
            new_title=effective_output.output_name,
            old_project_title=old_project_title,
            new_project_title=standard_project_title,
            title_replace_count=title_replace_count,
            project_title_replace_count=project_title_replace_count,
            document_number_replace_count=document_number_replace_count,
        )

        return {
            "status": "updated",
            "output_id": effective_output.output_id,
            "output_name": effective_output.output_name,
            "old_path": str(original_path),
            "new_path": str(final_path),
            "expected_filename": expected_filename,
            "file_name_changed": original_path.name != final_path.name,
            "backup_path": str(backup_path) if backup_path else "",
            **cover_status,
        }
    except Exception as exc:
        return {
            "status": "error",
            "output_id": candidate.output.output_id,
            "output_name": candidate.output.output_name,
            "old_path": str(original_path),
            "backup_path": str(backup_path) if backup_path else "",
            "error": str(exc),
        }


def build_cover_status(
    *,
    old_document_number: str,
    new_document_number: str,
    old_title: str | None,
    new_title: str,
    old_project_title: str | None,
    new_project_title: str,
    title_replace_count: int,
    project_title_replace_count: int,
    document_number_replace_count: int,
) -> dict[str, object]:
    document_number_changed = bool(old_document_number and old_document_number != new_document_number)
    if document_number_replace_count:
        document_number_changed = True

    title_change_expected = values_differ(old_title, new_title)
    project_title_change_expected = values_differ(old_project_title, new_project_title)
    title_changed = title_replace_count > 0
    project_title_changed = project_title_replace_count > 0
    cover_changed = document_number_changed or title_changed or project_title_changed

    warnings: list[str] = []
    if title_change_expected and not title_changed:
        warnings.append("문서명 변경 대상 텍스트를 찾지 못했습니다.")
    if project_title_change_expected and not project_title_changed:
        warnings.append("사업명 변경 대상 텍스트를 찾지 못했습니다.")
    if not cover_changed:
        warnings.append("표지 변경이 감지되지 않았습니다.")

    return {
        "cover_changed": cover_changed,
        "cover_warning_reasons": warnings,
    }


def values_differ(old_value: str | None, new_value: str | None) -> bool:
    if not old_value or not new_value:
        return False
    return normalize_for_match(old_value) != normalize_for_match(new_value)


def is_filename_unchanged_item(item: dict[str, object]) -> bool:
    if item.get("status") != "updated":
        return False
    return not bool(item.get("file_name_changed"))


def is_cover_unchanged_item(item: dict[str, object]) -> bool:
    if item.get("status") != "updated":
        return False
    return not bool(item.get("cover_changed"))


def is_cover_attention_item(item: dict[str, object]) -> bool:
    if item.get("status") == "error":
        return True
    if item.get("status") != "updated":
        return False
    return is_cover_unchanged_item(item) or bool(item.get("cover_warning_reasons"))


def write_apply_readme(
    dump_root: Path,
    standard_file: Path,
    standard_project_title: str,
    apply_items: list[dict[str, object]],
    skipped_file_count: int,
    *,
    requirement_result: RequirementGenerationResult | None = None,
) -> Path:
    report_path = unique_file_path(dump_root / "README_검수결과.md")
    report_path.write_text(
        build_apply_readme(
            dump_root,
            standard_file,
            standard_project_title,
            apply_items,
            skipped_file_count,
            requirement_result=requirement_result,
        ),
        encoding="utf-8",
    )
    return report_path


def build_apply_readme(
    dump_root: Path,
    standard_file: Path,
    standard_project_title: str,
    apply_items: list[dict[str, object]],
    skipped_file_count: int,
    *,
    requirement_result: RequirementGenerationResult | None = None,
) -> str:
    updated_items = [item for item in apply_items if item.get("status") == "updated"]
    failed_items = [item for item in apply_items if item.get("status") == "error"]
    filename_unchanged_items = [item for item in apply_items if is_filename_unchanged_item(item)]
    cover_attention_items = [item for item in apply_items if is_cover_attention_item(item)]

    lines = [
        "# 반영 결과 검수 README",
        "",
        "## 요약",
        "",
        "| 항목 | 값 |",
        "| --- | --- |",
        f"| 생성 시각 | {markdown_cell(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))} |",
        f"| 문서관리표준 | {markdown_cell(standard_file.name)} |",
        f"| 결과 폴더 | {markdown_cell(str(dump_root))} |",
        f"| 기준 사업명 | {markdown_cell(standard_project_title or '-')} |",
        f"| 반영 대상 | {len(apply_items)}건 |",
        f"| 반영 성공 | {len(updated_items)}건 |",
        f"| 반영 오류 | {len(failed_items)}건 |",
        f"| 제외/건너뜀 | {skipped_file_count}건 |",
        f"| 1차 파일명 미변경 | {len(filename_unchanged_items)}건 |",
        f"| 표지 확인 필요 파일 | {len(cover_attention_items)}건 |",
        "",
        "## 1차 파일명 미변경",
        "",
    ]

    append_filename_unchanged_section(lines, filename_unchanged_items, dump_root)
    lines.extend([
        "",
        "## 표지 확인 필요 파일",
        "",
        "반영은 끝났지만 표지의 문서번호, 산출물명, 사업명 변경 여부를 한 번 더 확인해야 하는 파일입니다.",
        "아래 표에서 파일 경로와 확인 내용을 보고 표지를 직접 확인하세요.",
        "",
    ])
    append_cover_attention_section(lines, cover_attention_items, dump_root)
    lines.extend(["", "## 반영 오류", ""])
    append_error_section(lines, failed_items, dump_root)
    append_requirement_generation_section(lines, requirement_result, dump_root)
    lines.append("")
    return "\n".join(lines)


def append_requirement_generation_section(
    lines: list[str],
    result: RequirementGenerationResult | None,
    dump_root: Path,
) -> None:
    if result is None or not result.enabled:
        return

    created_ok = [item for item in result.created_items if item.get("status") != "error"]
    warnings = [item for item in result.created_items if item.get("status") == "created_with_warning"]
    created_errors = [item for item in result.created_items if item.get("status") == "error"]
    removed_items = result.removed_items or []
    removed_ok = [item for item in removed_items if item.get("status") == "removed"]
    removed_errors = [item for item in removed_items if item.get("status") == "error"]
    folder_items = result.folder_items or []
    folder_ok = [item for item in folder_items if item.get("status") != "error"]
    folder_created = [item for item in folder_items if item.get("status") == "created"]
    folder_errors = [item for item in folder_items if item.get("status") == "error"]

    lines.extend(
        [
            "",
            "## 요구사항별 자동 생성",
            "",
            "| 항목 | 값 |",
            "| --- | --- |",
            f"| 대상 산출물 | {result.target_count}건 |",
            f"| 요구사항 ID 폴더 | {len(folder_ok)}건 (신규 {len(folder_created)}건) |",
            f"| 생성 파일 | {len(created_ok)}건 |",
            f"| 기존 기준 파일 삭제 | {len(removed_ok)}건 |",
            f"| 생성 경고 | {len(warnings)}건 |",
            f"| 요구사항 ID 없음 | {len(result.skipped_items)}건 |",
            f"| 생성 오류 | {len(created_errors) + len(result.error_items) + len(removed_errors) + len(folder_errors)}건 |",
            f"| 상세 README | {markdown_cell(relative_report_path(str(result.readme_path), dump_root) if result.readme_path else '-')} |",
        ]
    )


def append_filename_unchanged_section(
    lines: list[str],
    items: list[dict[str, object]],
    dump_root: Path,
) -> None:
    if not items:
        lines.append("없음")
        return

    lines.extend([
        "| 산출물ID | 산출물명 | 파일 | 예상 파일명 |",
        "| --- | --- | --- | --- |",
    ])
    for item in items:
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(item.get("output_id", "")),
                    markdown_cell(item.get("output_name", "")),
                    markdown_cell(item_report_path(item, dump_root)),
                    markdown_cell(item.get("expected_filename", "")),
                ]
            )
            + " |"
        )


def append_cover_attention_section(
    lines: list[str],
    items: list[dict[str, object]],
    dump_root: Path,
) -> None:
    if not items:
        lines.append("없음")
        return

    lines.extend([
        "| 상태 | 산출물ID | 산출물명 | 파일 | 확인 내용 |",
        "| --- | --- | --- | --- | --- |",
    ])
    for item in items:
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(cover_attention_label(item)),
                    markdown_cell(item.get("output_id", "")),
                    markdown_cell(item.get("output_name", "")),
                    markdown_cell(item_report_path(item, dump_root)),
                    markdown_cell(cover_attention_detail(item)),
                ]
            )
            + " |"
        )


def append_error_section(
    lines: list[str],
    items: list[dict[str, object]],
    dump_root: Path,
) -> None:
    if not items:
        lines.append("없음")
        return

    lines.extend([
        "| 산출물ID | 산출물명 | 파일 | 오류 |",
        "| --- | --- | --- | --- |",
    ])
    for item in items:
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(item.get("output_id", "")),
                    markdown_cell(item.get("output_name", "")),
                    markdown_cell(item_report_path(item, dump_root)),
                    markdown_cell(item.get("error", "")),
                ]
            )
            + " |"
        )


def cover_attention_label(item: dict[str, object]) -> str:
    if item.get("status") == "error":
        return "오류"
    if is_cover_unchanged_item(item):
        return "표지 미변경"
    return "부분 확인"


def cover_attention_detail(item: dict[str, object]) -> str:
    if item.get("status") == "error":
        return str(item.get("error", ""))
    warnings = item.get("cover_warning_reasons") or []
    if isinstance(warnings, list) and warnings:
        return " / ".join(str(value) for value in warnings)
    return "확인 필요"


def item_report_path(item: dict[str, object], dump_root: Path) -> str:
    path = str(item.get("new_path") or item.get("old_path") or "")
    return relative_report_path(path, dump_root)


def relative_report_path(path_text: str, dump_root: Path) -> str:
    if not path_text:
        return ""
    path = Path(path_text)
    try:
        return str(path.relative_to(dump_root))
    except ValueError:
        return path_text


def markdown_cell(value: object) -> str:
    text = str(value) if value is not None else ""
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    return text.replace("|", "\\|") or "-"
