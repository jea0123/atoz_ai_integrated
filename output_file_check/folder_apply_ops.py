# 덤프 폴더 복사, 백업, 문서값 반영, 파일명 변경을 실행합니다.
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import shutil
from uuid import uuid4

from app_runtime import (
    RESULT_DIR,
    log_event,
)
from document_update.document_number import write_updated_document
from document_update.patterns import OUTPUT_ID_PATTERN
from document_update.runtime_conversion import prepare_target_file
from output_file_check.folder_mapping import (
    build_folder_mapping,
    normalize_relative_path_for_compare,
    split_excluded_paths_field,
)
from output_file_check.folder_policy import FolderPolicy
from output_file_check.folder_serialization import serialize_check_result
from output_file_check.models import MatchCandidate, OutputMatch, StandardOutput
from output_file_check.normalization import filesystem_safe_stem, normalize_for_match, output_name_from_id


IGNORED_DUMP_FOLDER_NAMES = {"bak", "backup", "백업"}
TEST_OUTPUT_ID_PATTERN = re.compile(r"\d{4,}(?:\([^)]*\))?")
PRESERVED_TAIL_PATTERN = re.compile(
    r"((?:[_-]SFR-[A-Za-z0-9-]+)?(?:[_-][vV]\d+(?:\.\d+)*)?)$",
    re.IGNORECASE,
)


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
    preserved_tail = extract_preserved_tail(old_path.stem)
    stem = build_standard_stem(output)
    return f"{filesystem_safe_stem(stem)}{preserved_tail}{suffix}"


def build_standard_stem(output: StandardOutput) -> str:
    id_name = output_name_from_id(output.output_id)
    output_name = output.output_name

    if id_name and normalize_for_match(id_name) == normalize_for_match(output_name):
        return output.output_id

    if id_name:
        return output.output_id

    return f"{output.output_id}-{output_name}" if output.output_id else output_name


def extract_preserved_tail(stem: str) -> str:
    match = PRESERVED_TAIL_PATTERN.search(stem)
    return match.group(1) if match else ""


def copy_folder_to_dump(source_root: Path, dump_parent: Path) -> Path:
    # 원본 폴더를 건드리지 않기 위해 덤프 위치에 복사본을 만든다.
    dump_root = next_versioned_dump_path(dump_parent, source_root.name)
    shutil.copytree(source_root, dump_root, ignore=ignore_dump_backup_dirs)
    return dump_root


def ignore_dump_backup_dirs(directory: str, names: list[str]) -> set[str]:
    # 기존 백업 폴더는 결과물에 필요 없고 잠금/권한 오류 원인이 되므로 덤프에 복사하지 않는다.
    return {
        name for name in names
        if (Path(directory) / name).is_dir() and name.casefold() in IGNORED_DUMP_FOLDER_NAMES
    }


def apply_dumped_folder(
    standard_file: Path,
    dump_root: Path,
    fields: dict[str, str],
    folder_policy: FolderPolicy,
    temp_dir: Path,
    request_id: str,
    *,
    log_prefix: str,
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
    apply_items = [
        apply_batch_candidate(
            candidate,
            mapping.standard_project_title,
            temp_dir,
            rename_files=True,
        )
        for candidate in selected_candidates
    ]

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
            "apply_items": apply_items,
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


def resolve_dump_parent(raw_path: str, request_id: str) -> Path:
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
    backup_path = backup_original_for_batch(original_path)

    try:
        target_file, converted_to_hwpx = prepare_target_file(original_path, temp_dir)
        update_dir = temp_dir / "batch-updated"
        update_dir.mkdir(parents=True, exist_ok=True)
        output_suffix = target_file.suffix or original_path.suffix
        temp_output = unique_file_path(update_dir / f"{uuid4().hex}{output_suffix}")
        old_title = identity.document_title if identity and identity.document_title else None
        old_project_title = identity.project_title if identity and identity.project_title else None
        output_id = clean_output_id(candidate.output.output_id)
        validate_output_id(output_id)

        (
            _old_document_number,
            _document_backup_path,
            title_replace_count,
            project_title_replace_count,
            document_number_replace_count,
            output_file,
        ) = write_updated_document(
            target_file,
            new_document_number=output_id,
            old_title=old_title,
            new_title=candidate.output.output_name if old_title else None,
            old_project_title=old_project_title,
            new_project_title=standard_project_title,
            output_path=temp_output,
        )

        final_path = original_path if original_path.suffix.lower() == output_suffix.lower() else original_path.with_suffix(output_suffix)
        if rename_files:
            final_path = final_path.with_name(build_target_filename(candidate.output, final_path))
        if final_path.exists() and final_path.resolve() != original_path.resolve():
            final_path = unique_file_path(final_path)

        final_path.parent.mkdir(parents=True, exist_ok=True)
        output_file.replace(final_path)
        if original_path.exists() and original_path.resolve() != final_path.resolve():
            original_path.unlink()

        return {
            "status": "updated",
            "output_id": candidate.output.output_id,
            "output_name": candidate.output.output_name,
            "old_path": str(original_path),
            "new_path": str(final_path),
            "backup_path": str(backup_path) if backup_path else "",
            "converted_to_hwpx": converted_to_hwpx,
            "title_replace_count": title_replace_count,
            "project_title_replace_count": project_title_replace_count,
            "document_number_replace_count": document_number_replace_count,
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


def backup_original_for_batch(file_path: Path) -> Path | None:
    # 수정 전 원본 파일을 같은 폴더의 backup/bak 계열 폴더에 저장한다.
    backup_dir = find_existing_backup_dir(file_path.parent)
    if backup_dir is None:
        return None

    target = unique_file_path(backup_dir / file_path.name)
    shutil.copy2(file_path, target)
    return target


def find_existing_backup_dir(parent: Path) -> Path | None:
    # 이미 존재하는 원본/backup/bak 폴더를 찾는다.
    original_dir = find_child_dir(parent, {"원본"})
    if original_dir:
        return original_dir
    return find_child_dir(parent, {"bak", "backup", "백업"})


def find_child_dir(parent: Path, names: set[str]) -> Path | None:
    # 대소문자를 무시하고 지정한 이름의 하위 폴더를 찾는다.
    normalized = {name.casefold() for name in names}
    for child in parent.iterdir():
        if child.is_dir() and child.name.casefold() in normalized:
            return child
    return None
