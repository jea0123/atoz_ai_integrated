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
from document_update.document_number import write_updated_document, write_updated_project_title
from document_update.hwp_convert import start_allow_all_watcher, stop_allow_all_watcher
from document_update.metadata_update import update_metadata_in_document
from document_update.patterns import OUTPUT_ID_PATTERN
from document_update.project_title_match import project_title_matches_expected
from document_update.runtime_conversion import prepare_target_file
from output_file_check.file_noise import copytree_ignore_noise, remove_noise_files
from output_file_check.content_identity import find_matching_project_title, read_file_identity
from output_file_check.folder_mapping import (
    build_folder_mapping,
    normalize_relative_path_for_compare,
    split_excluded_paths_field,
)
from output_file_check.folder_policy import FolderPolicy
from output_file_check.folder_scanner import scan_folder
from output_file_check.folder_serialization import serialize_check_result
from output_file_check.matcher import score_file
from output_file_check.models import MatchCandidate, OutputMatch, PathTemplate, ScannedFile, StandardOutput
from output_file_check.normalization import filesystem_safe_stem, normalize_for_match, output_name_from_id
from output_file_check.requirement_generation import (
    RequirementGenerationResult,
    generate_requirement_documents,
)


PRESERVED_TAIL_PATTERN = re.compile(
    r"((?:(?:[_-][vV]\d+(?:\.\d+)*)|(?:[_-]SFR-[A-Za-z0-9-]+))+)$",
    re.IGNORECASE,
)
UNMATCHED_FULL_APPLY_THRESHOLD = 0.94
ATTACHMENT_TAIL_PATTERN = re.compile(
    r"^[\s_-]*(?:[\[\(（［｛]\s*(?:별첨|첨부)\s*\d*[^)\]\}）］｝]*[\)\]\}）］｝]|(?:별첨|첨부)\s*\d+)",
    re.IGNORECASE,
)
ATTACHMENT_TAIL_SEARCH_PATTERN = re.compile(
    r"(?P<tail>[\s_-]*(?:[\[\(（［｛]\s*(?:별첨|첨부)\s*\d*[^)\]\}）］｝]*[\)\]\}）］｝]|(?:별첨|첨부)\s*\d+).*)",
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
    # 표준의 산출물 ID는 사용자가 프로젝트마다 바꿔 넣는 값이다.
    # MFDS-* 형식만 허용하면 관리산출물의 "11", "112233" 같은 값이 적용 전에 전부 실패한다.
    if not output_id:
        raise RuntimeError("산출물 ID가 비어 있습니다.")
    if re.search(r"[\x00-\x1f\x7f]", output_id):
        raise RuntimeError(f"산출물 ID에 사용할 수 없는 제어문자가 있습니다: {output_id}")


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
        attachment_tail = attachment_tail_from_stem(stem)
        if attachment_tail:
            tail = attachment_tail
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


def attachment_tail_from_stem(stem: str) -> str:
    match = ATTACHMENT_TAIL_SEARCH_PATTERN.search(stem)
    return match.group("tail").strip(" \t\r\n") if match else ""


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
    revision_metadata = initial_revision_metadata_from_fields(fields)

    all_selected_candidates = select_all_unique_candidates(mapping.matches)
    excluded_relative_paths = split_excluded_paths_field(fields.get("excluded_candidate_paths"))
    selected_candidates = filter_excluded_candidates(
        all_selected_candidates,
        dump_root,
        excluded_relative_paths,
    )
    selected_candidates = filter_apply_scope_candidates(
        selected_candidates,
        dump_root,
        fields,
    )
    excluded_file_count = len(all_selected_candidates) - len(selected_candidates)
    allow_stop_event, allow_thread = start_allow_all_watcher()
    try:
        apply_items = [
            apply_batch_candidate(
                candidate,
                mapping.standard_project_title,
                temp_dir,
                revision_metadata=revision_metadata,
                rename_files=True,
            )
            for candidate in selected_candidates
        ]
        apply_items.extend(
            apply_unmatched_project_title_updates(
                dump_root,
                mapping.outputs,
                mapping.files,
                selected_candidates,
                mapping.standard_project_title,
                temp_dir,
                excluded_relative_paths,
                revision_metadata=revision_metadata,
                ignore_top_level_files=is_management_apply_scope(fields),
            )
        )
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
            "project_title_only_updated_count": sum(
                1 for item in apply_items
                if item.get("status") == "updated"
                and item.get("project_only")
                and int(item.get("cover_project_replace_count") or 0) > 0
            ),
            "skipped_file_count": excluded_file_count,
            "filename_unchanged_count": sum(1 for item in apply_items if is_filename_unchanged_item(item)),
            "initial_revision_date": revision_metadata["revision_date"],
            "initial_revision_author": revision_metadata["author"],
            "initial_revision_approval_author": revision_metadata["approval_author"],
            "initial_revision_updated_count": sum(
                1 for item in apply_items if item.get("initial_revision_status") == "updated"
            ),
            "initial_revision_skipped_count": sum(
                1 for item in apply_items if item.get("initial_revision_status") == "skipped"
            ),
            "initial_revision_failed_count": sum(
                1 for item in apply_items if item.get("initial_revision_status") == "error"
            ),
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


def filter_apply_scope_candidates(
    candidates: list[MatchCandidate],
    dump_root: Path,
    fields: dict[str, str],
) -> list[MatchCandidate]:
    if not is_management_apply_scope(fields):
        return candidates
    return [
        candidate
        for candidate in candidates
        if not is_top_level_dump_file(candidate.file.path, dump_root)
    ]


def is_management_apply_scope(fields: dict[str, str]) -> bool:
    return str(fields.get("artifact_category") or "").strip().casefold() == "management"


def is_top_level_dump_file(path: Path, dump_root: Path) -> bool:
    try:
        return len(path.relative_to(dump_root).parts) == 1
    except ValueError:
        return False


def apply_unmatched_project_title_updates(
    dump_root: Path,
    outputs: list[StandardOutput],
    files: list[ScannedFile],
    selected_candidates: list[MatchCandidate],
    standard_project_title: str,
    temp_dir: Path,
    excluded_relative_paths: set[str],
    *,
    revision_metadata: dict[str, str],
    ignore_top_level_files: bool = False,
) -> list[dict[str, object]]:
    selected_paths = selected_candidate_path_keys(selected_candidates)
    items: list[dict[str, object]] = []
    seen: set[str] = set()

    for file in project_title_update_candidates(dump_root, files):
        path = file.path
        path_key = str(path.resolve(strict=False)).casefold()
        if path_key in selected_paths or path_key in seen:
            continue
        seen.add(path_key)
        if ignore_top_level_files and is_top_level_dump_file(path, dump_root):
            continue
        if not path.exists() or not path.is_file():
            continue
        try:
            relative_path = str(path.relative_to(dump_root))
        except ValueError:
            relative_path = str(path)
        if normalize_relative_path_for_compare(relative_path) in excluded_relative_paths:
            continue

        identity = file.identity
        file = ScannedFile(path, identity)
        output_candidate = best_unselected_output_candidate(file, outputs)
        if output_candidate is not None:
            items.append(
                apply_batch_candidate(
                    output_candidate,
                    standard_project_title,
                    temp_dir,
                    revision_metadata=revision_metadata,
                    rename_files=True,
                )
            )
            continue

        if identity is None:
            if should_skip_expensive_unmatched_identity_read(path):
                continue
            identity = read_file_identity(path)
            file = ScannedFile(path, identity)

        old_project_title = project_title_for_update(identity, standard_project_title)
        if not old_project_title or not values_differ(old_project_title, standard_project_title):
            continue

        items.append(
            apply_project_title_only_file(
                path,
                old_project_title,
                standard_project_title,
                temp_dir,
                revision_metadata,
            )
        )

    return items


def selected_candidate_path_keys(selected_candidates: list[MatchCandidate]) -> set[str]:
    keys: set[str] = set()
    for candidate in selected_candidates:
        original_path = candidate.file.path
        paths = [original_path, original_path.with_name(build_target_filename(candidate.output, original_path))]
        if original_path.suffix.lower() == ".hwp":
            converted_path = original_path.with_suffix(".hwpx")
            paths.append(converted_path)
            paths.append(converted_path.with_name(build_target_filename(candidate.output, converted_path)))
        keys.update(str(path.resolve(strict=False)).casefold() for path in paths)
    return keys


def best_unselected_output_candidate(
    file: ScannedFile,
    outputs: list[StandardOutput],
) -> MatchCandidate | None:
    best: MatchCandidate | None = None
    for output in outputs:
        candidate = score_file(output, file, use_output_id=True)
        if candidate is None:
            continue
        if best is None or candidate.score > best.score:
            best = candidate
    if best is None or best.score < UNMATCHED_FULL_APPLY_THRESHOLD:
        return None
    return best


def project_title_update_candidates(dump_root: Path, files: list[ScannedFile]) -> list[ScannedFile]:
    candidates: list[ScannedFile] = []
    seen: set[str] = set()
    for file in files:
        key = str(file.path.resolve(strict=False)).casefold()
        if key in seen:
            continue
        candidates.append(file)
        seen.add(key)

    for file in scan_folder(dump_root, read_contents=False, folder_policy=None):
        key = str(file.path.resolve(strict=False)).casefold()
        if key in seen or has_ignored_project_title_path(file.path, dump_root):
            continue
        candidates.append(file)
        seen.add(key)
    return candidates


def should_skip_expensive_unmatched_identity_read(path: Path) -> bool:
    return path.suffix.lower() in {".hwp", ".hwpx"}


def has_ignored_project_title_path(path: Path, dump_root: Path) -> bool:
    try:
        parts = path.relative_to(dump_root).parts[:-1]
    except ValueError:
        parts = path.parts[:-1]
    ignored = {normalize_for_match(value) for value in ("bak", "backup", "백업", "임시", "temp", "tmp")}
    return any(normalize_for_match(part) in ignored for part in parts)


def apply_project_title_only_file(
    original_path: Path,
    old_project_title: str,
    standard_project_title: str,
    temp_dir: Path,
    revision_metadata: dict[str, str],
) -> dict[str, object]:
    try:
        target_file, _converted_to_hwpx = prepare_target_file(original_path, temp_dir)
        update_dir = temp_dir / "project-title-updated"
        update_dir.mkdir(parents=True, exist_ok=True)
        output_suffix = target_file.suffix or original_path.suffix
        temp_output = unique_file_path(update_dir / f"{uuid4().hex}{output_suffix}")
        replace_count, output_file = write_updated_project_title(
            target_file,
            old_project_title,
            standard_project_title,
            output_path=temp_output,
        )
        if replace_count <= 0:
            return {
                "status": "updated",
                "project_only": True,
                "output_id": "",
                "output_name": "표준 외 문서",
                "old_path": str(original_path),
                "new_path": str(original_path),
                "expected_filename": "",
                "file_name_changed": False,
                "cover_changed": False,
                "cover_update_status": "unchanged",
                "cover_project_replace_count": 0,
                "cover_document_number_replace_count": 0,
                "cover_update_error": "",
                "cover_warning_reasons": ["표지에서 기존 사업명 텍스트를 교체하지 못했습니다."],
            }

        final_path = original_path if original_path.suffix.lower() == output_suffix.lower() else original_path.with_suffix(output_suffix)
        if not same_path(output_file, final_path):
            replace_file_with_fallback(output_file, final_path)
        file_cleanup_error = cleanup_original_after_replacement(original_path, final_path)

        revision_result = apply_initial_revision_metadata(final_path, temp_dir, revision_metadata)

        return {
            "status": "updated",
            "project_only": True,
            "output_id": "",
            "output_name": "표준 외 문서",
            "old_path": str(original_path),
            "new_path": str(final_path),
            "expected_filename": "",
            "file_name_changed": original_path.name != final_path.name,
            "cover_changed": True,
            "cover_update_status": "updated",
            "cover_project_replace_count": replace_count,
            "cover_document_number_replace_count": 0,
            "cover_update_error": "",
            "cover_warning_reasons": [],
            "file_cleanup_error": file_cleanup_error,
            "initial_revision_status": revision_result["status"],
            "initial_revision_date": revision_result["revision_date"],
            "initial_revision_author": revision_result["author"],
            "initial_revision_approval_author": revision_result["approval_author"],
            "initial_revision_cover_update_count": revision_result["cover_update_count"],
            "initial_revision_history_update_count": revision_result["revision_history_update_count"],
            "initial_revision_error": revision_result["error"],
        }
    except Exception as exc:
        return {
            "status": "error",
            "project_only": True,
            "output_id": "",
            "output_name": "표준 외 문서",
            "old_path": str(original_path),
            "backup_path": "",
            "error": f"프로젝트명만 교체 실패: {exc}",
        }


def apply_batch_candidate(
    candidate: MatchCandidate,
    standard_project_title: str,
    temp_dir: Path,
    *,
    revision_metadata: dict[str, str],
    rename_files: bool,
) -> dict[str, object]:
    # 후보 파일 하나에 산출물 ID/제목/프로젝트명을 반영하고 필요하면 파일명도 바꾼다.
    original_path = candidate.file.path
    identity = candidate.file.identity
    backup_path = None

    try:
        target_file, _converted_to_hwpx = prepare_target_file(original_path, temp_dir)
        update_dir = temp_dir / "batch-updated"
        update_dir.mkdir(parents=True, exist_ok=True)
        output_suffix = target_file.suffix or original_path.suffix
        temp_output = unique_file_path(update_dir / f"{uuid4().hex}{output_suffix}")
        effective_output = candidate.output
        old_project_title = project_title_for_update(
            identity,
            standard_project_title,
            effective_output,
        )
        output_id = clean_output_id(effective_output.output_id)
        validate_output_id(output_id)

        old_document_number = ""
        project_title_replace_count = 0
        document_number_replace_count = 0
        cover_update_error = ""
        output_file = target_file
        try:
            (
                old_document_number,
                _document_backup_path,
                project_title_replace_count,
                document_number_replace_count,
                output_file,
            ) = write_updated_document(
                target_file,
                new_document_number=output_id,
                old_project_title=old_project_title,
                new_project_title=standard_project_title,
                output_path=temp_output,
                allow_missing_document_number=True,
            )
        except Exception as exc:
            # 표지 형식이 달라도 파일명 정리는 계속 진행하고, 결과 README에 확인 내용을 남긴다.
            cover_update_error = str(exc)

        final_path = original_path if original_path.suffix.lower() == output_suffix.lower() else original_path.with_suffix(output_suffix)
        expected_filename = build_target_filename(effective_output, final_path)
        if rename_files:
            final_path = final_path.with_name(expected_filename)
        if final_path.exists() and not same_path(final_path, original_path):
            final_path = unique_file_path(final_path)

        final_path.parent.mkdir(parents=True, exist_ok=True)
        if not same_path(output_file, final_path):
            replace_file_with_fallback(output_file, final_path)
        file_cleanup_error = cleanup_original_after_replacement(original_path, final_path)

        revision_result = apply_initial_revision_metadata(
            final_path,
            temp_dir,
            revision_metadata,
            document_number=output_id,
        )
        project_title_verified = project_title_verified_in_file(final_path, standard_project_title)

        cover_status = build_cover_status(
            old_document_number=old_document_number,
            new_document_number=output_id,
            old_project_title=old_project_title,
            new_project_title=standard_project_title,
            project_title_replace_count=project_title_replace_count,
            document_number_replace_count=document_number_replace_count,
            cover_update_error=cover_update_error,
            project_title_verified=project_title_verified,
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
            "file_cleanup_error": file_cleanup_error,
            "initial_revision_status": revision_result["status"],
            "initial_revision_date": revision_result["revision_date"],
            "initial_revision_author": revision_result["author"],
            "initial_revision_approval_author": revision_result["approval_author"],
            "initial_revision_cover_update_count": revision_result["cover_update_count"],
            "initial_revision_history_update_count": revision_result["revision_history_update_count"],
            "initial_revision_error": revision_result["error"],
            "cover_project_replace_count": project_title_replace_count,
            "cover_project_title_verified": project_title_verified,
            "cover_document_number_replace_count": document_number_replace_count,
            "cover_update_error": cover_update_error,
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


def same_path(left: Path, right: Path) -> bool:
    return left.resolve(strict=False) == right.resolve(strict=False)


def replace_file_with_fallback(source: Path, target: Path) -> None:
    try:
        source.replace(target)
        return
    except OSError:
        shutil.copyfile(source, target)
        try:
            source.unlink()
        except OSError:
            pass


def cleanup_original_after_replacement(original_path: Path, final_path: Path) -> str:
    if same_path(original_path, final_path) or not original_path.exists():
        return ""
    try:
        original_path.unlink()
        return ""
    except OSError as exc:
        return str(exc)


def project_title_for_update(
    identity: FileIdentity | None,
    standard_project_title: str,
    output: StandardOutput | None = None,
) -> str | None:
    if identity is None:
        return None

    candidates: list[str] = []
    if identity.project_title:
        candidates.append(identity.project_title)
    if identity.document_title:
        candidates.append(identity.document_title)

    matched = find_matching_project_title(
        identity.preview_text or "",
        standard_project_title,
        tuple(candidates),
    )
    if matched and (output is None or normalized_project_title_candidate(matched, output)):
        return matched

    fallback = normalized_project_title_candidate(identity.project_title, output) if output else str(identity.project_title or "").strip()
    if not fallback:
        return None

    matched = find_matching_project_title(
        "\n".join(value for value in (identity.preview_text, fallback) if value),
        standard_project_title,
        (fallback,),
    )
    return matched or None


def normalized_project_title_candidate(value: str, output: StandardOutput) -> str | None:
    text = str(value or "").strip()
    if not text:
        return None

    text_key = normalize_for_match(text)
    output_keys = [
        normalize_for_match(output.output_name),
        *(normalize_for_match(alias) for alias in output.aliases),
    ]
    if any(output_key and text_key == output_key for output_key in output_keys):
        return None
    return text


def initial_revision_year(value: object = "") -> str:
    text = str(value or "").strip()
    if not text:
        return str(datetime.now().year)
    if not re.fullmatch(r"\d{4}", text):
        raise ValueError("개정일자는 연도 4자리로 입력하세요.")
    return text


def initial_revision_date(year: object = "") -> str:
    return f"{initial_revision_year(year)}-00-00"


def initial_revision_metadata_from_fields(fields: dict[str, str]) -> dict[str, str]:
    return {
        "revision_date": initial_revision_date(fields.get("initial_revision_year", "")),
        "author": (fields.get("initial_revision_author") or "송아름").strip() or "송아름",
        "approval_author": (fields.get("initial_revision_approval_author") or "임채현").strip() or "임채현",
    }


def apply_initial_revision_metadata(
    file_path: Path,
    temp_dir: Path,
    revision_metadata: dict[str, str],
    document_number: str = "",
) -> dict[str, object]:
    revision_date = revision_metadata["revision_date"]
    author = revision_metadata["author"]
    approval_author = revision_metadata["approval_author"]
    try:
        result = update_metadata_in_document(
            file_path,
            author,
            revision_date,
            approval_author,
            temp_dir,
            document_number=document_number,
        )
        return {
            "status": result.status,
            "revision_date": revision_date,
            "author": author,
            "approval_author": approval_author,
            "cover_update_count": result.cover_update_count,
            "revision_history_update_count": result.revision_history_update_count,
            "error": result.error,
        }
    except Exception as exc:
        return {
            "status": "error",
            "revision_date": revision_date,
            "author": author,
            "approval_author": approval_author,
            "cover_update_count": 0,
            "revision_history_update_count": 0,
            "error": str(exc),
        }


def build_cover_status(
    *,
    old_document_number: str,
    new_document_number: str,
    old_project_title: str | None,
    new_project_title: str,
    project_title_replace_count: int,
    document_number_replace_count: int,
    cover_update_error: str = "",
    project_title_verified: bool = False,
) -> dict[str, object]:
    document_number_changed = bool(old_document_number and old_document_number != new_document_number)
    if document_number_replace_count:
        document_number_changed = True

    project_title_change_expected = values_differ(old_project_title, new_project_title)
    project_title_changed = project_title_replace_count > 0
    project_title_ok = project_title_changed or project_title_verified
    cover_changed = document_number_changed or project_title_changed
    document_number_missing = not old_document_number and document_number_replace_count == 0

    warnings: list[str] = []
    if cover_update_error:
        warnings.append(f"표지 수정 실패: {cover_update_error}")
    else:
        if document_number_missing:
            warnings.append("표지에서 문서번호 위치를 찾지 못해 내부 문서번호는 확인 필요")
        if project_title_change_expected and not project_title_ok:
            warnings.append("사업명 변경 대상 텍스트를 찾지 못했습니다.")
        elif new_project_title and not old_project_title and not project_title_ok:
            warnings.append("표지에서 기존 사업명 후보를 확정하지 못했습니다.")
        if not cover_changed and not project_title_verified:
            warnings.append("표지 변경이 감지되지 않았습니다.")

    if cover_update_error:
        cover_update_status = "failed"
    elif cover_changed and warnings:
        cover_update_status = "partial"
    elif cover_changed:
        cover_update_status = "updated"
    else:
        cover_update_status = "unchanged"

    return {
        "cover_changed": cover_changed,
        "cover_update_status": cover_update_status,
        "cover_document_number_missing": document_number_missing,
        "cover_warning_reasons": warnings,
    }


def project_title_verified_in_file(file_path: Path, expected_project_title: str) -> bool:
    if not expected_project_title:
        return False
    try:
        identity = read_file_identity(file_path)
    except Exception:
        return False
    if identity.error:
        return False
    if project_title_matches_expected(identity.project_title, expected_project_title):
        return True
    if expected_project_title in identity.preview_text:
        return True
    if normalize_for_match(expected_project_title) in normalize_for_match(identity.preview_text):
        return True
    matched = find_matching_project_title(
        identity.preview_text,
        expected_project_title,
        extra_candidates=(identity.project_title,),
    )
    return bool(matched)


def values_differ(old_value: str | None, new_value: str | None) -> bool:
    if not old_value or not new_value:
        return False
    return normalize_for_match(old_value) != normalize_for_match(new_value)


def is_filename_unchanged_item(item: dict[str, object]) -> bool:
    if item.get("status") != "updated":
        return False
    if item.get("project_only"):
        return False
    return not bool(item.get("file_name_changed"))


def is_cover_unchanged_item(item: dict[str, object]) -> bool:
    if item.get("status") != "updated":
        return False
    return not bool(item.get("cover_changed"))


def is_author_attention_item(item: dict[str, object]) -> bool:
    if item.get("status") != "updated":
        return False
    status = str(item.get("initial_revision_status") or "")
    if status in {"skipped", "error"}:
        return True
    if item.get("initial_revision_error"):
        return True
    cover_count = int(item.get("initial_revision_cover_update_count") or 0)
    history_count = int(item.get("initial_revision_history_update_count") or 0)
    return status == "updated" and cover_count == 0 and history_count == 0


def is_header_id_attention_item(item: dict[str, object]) -> bool:
    if item.get("status") != "updated" or item.get("project_only"):
        return False
    if item.get("cover_document_number_missing"):
        return True
    warnings = item.get("cover_warning_reasons") or []
    return isinstance(warnings, list) and any("문서번호" in str(value) for value in warnings)


def is_project_title_attention_item(item: dict[str, object]) -> bool:
    if item.get("status") != "updated":
        return False
    warnings = item.get("cover_warning_reasons") or []
    if isinstance(warnings, list) and any(
        ("사업명" in str(value) or "프로젝트명" in str(value))
        for value in warnings
    ):
        return True
    return bool(item.get("project_only")) and int(item.get("cover_project_replace_count") or 0) == 0


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
    author_attention_items = [item for item in apply_items if is_author_attention_item(item)]
    header_id_attention_items = [item for item in apply_items if is_header_id_attention_item(item)]
    project_title_attention_items = [item for item in apply_items if is_project_title_attention_item(item)]
    cover_project_items = [item for item in updated_items if int(item.get("cover_project_replace_count") or 0) > 0]
    cover_document_number_items = [
        item for item in updated_items if int(item.get("cover_document_number_replace_count") or 0) > 0
    ]
    cover_partial_items = [item for item in updated_items if item.get("cover_update_status") == "partial"]
    cover_failed_items = [item for item in updated_items if item.get("cover_update_status") == "failed"]
    project_only_items = [
        item for item in updated_items
        if item.get("project_only") and int(item.get("cover_project_replace_count") or 0) > 0
    ]

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
        f"| 표지 사업명 교체 | {len(cover_project_items)}건 |",
        f"| 표준 외 문서 사업명 교체 | {len(project_only_items)}건 |",
        f"| 표지 문서번호 교체 | {len(cover_document_number_items)}건 |",
        f"| 표지 부분반영 | {len(cover_partial_items)}건 |",
        f"| 표지 수정 실패 | {len(cover_failed_items)}건 |",
        f"| 작성자/개정정보 확인 필요 | {len(author_attention_items)}건 |",
        f"| 머릿말/표지 ID 확인 필요 | {len(header_id_attention_items)}건 |",
        f"| 사업명 확인 필요 | {len(project_title_attention_items)}건 |",
    ]

    lines.extend([
        "",
        "## 작성자/개정정보 확인 필요",
        "",
        "작성자, 개정일자, 개정이력 반영이 없거나 실패한 파일입니다.",
        "",
    ])
    append_attention_category_section(lines, author_attention_items, dump_root, revision_attention_label, revision_attention_detail)
    lines.extend([
        "",
        "## 머릿말/표지 ID 확인 필요",
        "",
        "문서번호 또는 머릿말 ID 위치를 확정하지 못한 파일입니다.",
        "",
    ])
    append_attention_category_section(lines, header_id_attention_items, dump_root, cover_attention_label, header_id_attention_detail)
    lines.extend([
        "",
        "## 사업명 확인 필요",
        "",
        "표지에서 기준 사업명으로 교체할 기존 사업명 위치를 찾지 못한 파일입니다.",
        "",
    ])
    append_attention_category_section(lines, project_title_attention_items, dump_root, cover_attention_label, project_title_attention_detail)
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


def append_attention_category_section(
    lines: list[str],
    items: list[dict[str, object]],
    dump_root: Path,
    label_func,
    detail_func,
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
                    markdown_cell(label_func(item)),
                    markdown_cell(item.get("output_id", "")),
                    markdown_cell(item.get("output_name", "")),
                    markdown_cell(item_report_path(item, dump_root)),
                    markdown_cell(detail_func(item)),
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
    cover_status = str(item.get("cover_update_status") or "")
    if cover_status == "failed":
        return "표지 수정 실패"
    if cover_status == "partial":
        return "표지 부분반영"
    if is_cover_unchanged_item(item):
        return "표지 미변경"
    return "부분 확인"


def revision_attention_label(item: dict[str, object]) -> str:
    status = str(item.get("initial_revision_status") or "확인 필요")
    if status == "updated":
        return "반영 없음"
    if status == "skipped":
        return "건너뜀"
    if status == "error":
        return "오류"
    return status


def revision_attention_detail(item: dict[str, object]) -> str:
    status = item.get("initial_revision_status")
    error = str(item.get("initial_revision_error") or "")
    if error:
        return error
    if status == "skipped":
        return "문서 내부에서 수정할 표지/개정이력 위치를 찾지 못했습니다."

    details: list[str] = []
    cover_count = int(item.get("initial_revision_cover_update_count") or 0)
    history_count = int(item.get("initial_revision_history_update_count") or 0)
    if cover_count > 0:
        details.append(f"표지/머릿말 {cover_count}곳")
    if history_count > 0:
        details.append(f"개정이력 {history_count}곳")
    return " / ".join(details) if details else "확인 필요"


def header_id_attention_detail(item: dict[str, object]) -> str:
    details: list[str] = []
    if item.get("cover_document_number_missing"):
        details.append("문서번호 위치 미확정")
    document_count = int(item.get("cover_document_number_replace_count") or 0)
    details.append(f"문서번호 교체 {document_count}곳")
    warnings = item.get("cover_warning_reasons") or []
    if isinstance(warnings, list):
        details.extend(str(value) for value in warnings if "문서번호" in str(value))
    return " / ".join(details) if details else "확인 필요"


def project_title_attention_detail(item: dict[str, object]) -> str:
    details: list[str] = []
    project_count = int(item.get("cover_project_replace_count") or 0)
    if project_count > 0:
        details.append(f"사업명 교체 {project_count}곳")
    warnings = item.get("cover_warning_reasons") or []
    if isinstance(warnings, list):
        details.extend(
            str(value)
            for value in warnings
            if "사업명" in str(value) or "프로젝트명" in str(value)
        )
    return " / ".join(details) if details else "확인 필요"


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
