# 덤프 폴더 복사, 백업, 문서값 반영, 파일명 변경을 실행합니다.
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
from output_file_check.folder_mapping import (
    build_folder_mapping,
    normalize_relative_path_for_compare,
    split_excluded_paths_field,
)
from output_file_check.folder_policy import FolderPolicy
from output_file_check.folder_serialization import serialize_check_result
from output_file_check.models import MatchCandidate, OutputMatch, StandardOutput
from output_file_check.normalization import filesystem_safe_stem, normalize_for_match, output_name_from_id


TEST_OUTPUT_ID_PATTERN = re.compile(r"\d{4,}(?:\([^)]*\))?")
PRESERVED_TAIL_PATTERN = re.compile(
    r"((?:(?:[_-][vV]\d+(?:\.\d+)*)|(?:[_-]SFR-[A-Za-z0-9-]+)|(?:\[[^\]]+\]))+)$",
    re.IGNORECASE,
)
VERSION_TAIL_PATTERN = re.compile(r"([_-])[vV]\d+(?:\.\d+)*")


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
    if not match:
        return ""
    return VERSION_TAIL_PATTERN.sub(r"\1v0.1", match.group(1))


def copy_folder_to_dump(source_root: Path, dump_parent: Path) -> Path:
    # 원본 폴더를 건드리지 않기 위해 덤프 위치에 복사본을 만든다.
    dump_root = next_versioned_dump_path(dump_parent, source_root.name)
    shutil.copytree(windows_long_path(source_root), windows_long_path(dump_root))
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

    try:
        write_apply_readme(
            dump_root,
            standard_file,
            mapping.standard_project_title,
            apply_items,
            excluded_file_count,
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
    backup_path = backup_original_for_batch(original_path)

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
) -> Path:
    report_path = unique_file_path(dump_root / "README_검수결과.md")
    report_path.write_text(
        build_apply_readme(
            dump_root,
            standard_file,
            standard_project_title,
            apply_items,
            skipped_file_count,
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
    lines.append("")
    return "\n".join(lines)


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
