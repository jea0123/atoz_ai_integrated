# 문서관리표준, 실제 폴더, 매칭 전략을 조립해 최종 매핑 결과를 만듭니다.
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from app_runtime import (
    read_runtime_env,
    resolve_runtime_model,
    resolve_runtime_ollama_generate_url,
    selected_match_mode,
)
from output_file_check.content_identity import read_file_identity, read_standard_project_title
from output_file_check.folder_matching import (
    build_outputs_from_path_templates,
    find_templates_for_file_path,
    match_files_by_ai_first,
    match_files_by_folder_path,
)
from output_file_check.folder_policy import FolderPolicy
from output_file_check.folder_scanner import scan_folder
from output_file_check.matcher import DEFAULT_MATCH_THRESHOLD
from output_file_check.models import OutputMatch, PathTemplate, ScannedFile, StandardOutput
from output_file_check.path_template_reader import read_path_templates
from output_file_check.standard_reader import read_standard_outputs


@dataclass(frozen=True)
class FolderMappingResult:
    standard_project_title: str
    outputs: list[StandardOutput]
    path_templates: list[PathTemplate]
    files: list[ScannedFile]
    matches: list[OutputMatch]
    match_mode: str = "rule"


def build_folder_policy_from_fields(fields: dict[str, str]) -> FolderPolicy:
    # 웹/CLI 입력값에서 제외 폴더, 투명 폴더, 제한 경로 정책을 만든다.
    ignore = split_policy_field(fields.get("ignore_folder_names", "bak,backup,백업,임시,temp,tmp"))
    transparent = split_policy_field(fields.get("transparent_folder_names", "원본"))
    map_only = [
        tuple(part.strip() for part in item.replace("/", "\\").split("\\") if part.strip())
        for item in split_policy_field(fields.get("map_only_under", ""))
    ]
    return FolderPolicy(
        ignore_folder_names=tuple(ignore),
        transparent_folder_names=tuple(transparent),
        map_only_under=tuple(path for path in map_only if path),
    )


def split_policy_field(value: str | None) -> list[str]:
    # 쉼표/세미콜론/줄바꿈으로 들어온 정책 입력값을 리스트로 쪼갠다.
    if not value:
        return []
    return [item.strip() for item in re.split(r"[,;\n]+", value) if item.strip()]


def split_excluded_paths_field(value: str | None) -> set[str]:
    # 사용자가 반영 제외로 체크한 상대 경로들을 비교용 set으로 만든다.
    return {normalize_relative_path_for_compare(item) for item in split_policy_field(value)}


def normalize_relative_path_for_compare(value: str) -> str:
    # Windows/브라우저 경로 표기 차이를 없애 비교 가능한 문자열로 만든다.
    return value.replace("/", "\\").strip().casefold()


def build_folder_mapping(
    standard_file: Path,
    folder_dir: Path,
    fields: dict[str, str],
    folder_policy: FolderPolicy,
) -> FolderMappingResult:
    # 표준 PDF 읽기, 실제 폴더 스캔, 매칭 전략 실행을 한 번에 조립한다.
    """표준 PDF와 실제 폴더를 읽어 화면/반영 공통 매칭 결과를 만든다."""
    # 웹 화면에는 project_title 입력칸이 없다. 이 값은 CLI에서 --project-title로
    # 표준 PDF 사업명 파싱 결과를 강제로 보정할 때만 들어온다.
    project_title_override = fields.get("project_title", "").strip()
    standard_project_title = project_title_override or read_standard_project_title(standard_file)
    threshold = float(fields.get("threshold") or DEFAULT_MATCH_THRESHOLD)
    standard_outputs = read_standard_outputs(standard_file)
    path_templates = read_path_templates(standard_file, standard_outputs)
    outputs = build_outputs_from_path_templates(standard_outputs, path_templates)
    runtime_env = read_runtime_env()
    ollama_url = resolve_runtime_ollama_generate_url(runtime_env)
    model = resolve_runtime_model(runtime_env)

    match_strategy, match_mode = selected_match_mode(fields, ollama_url)
    files = scan_template_files(
        folder_dir,
        path_templates,
        folder_policy,
        read_contents=match_strategy != "ai_first",
    )

    if match_strategy == "ai_first" and ollama_url:
        matches = match_files_by_ai_first(
            outputs,
            files,
            path_templates,
            folder_dir,
            threshold=threshold,
            expected_project_title=standard_project_title,
            folder_policy=folder_policy,
            ollama_url=ollama_url,
            model=model,
        )
    else:
        matches = match_files_by_folder_path(
            outputs,
            files,
        path_templates,
        folder_dir,
        threshold=threshold,
        folder_policy=folder_policy,
    )

    return FolderMappingResult(
        standard_project_title=standard_project_title,
        outputs=outputs,
        path_templates=path_templates,
        files=files,
        matches=matches,
        match_mode=match_mode,
    )


def scan_template_files(
    folder_dir: Path,
    path_templates: list[PathTemplate],
    folder_policy: FolderPolicy,
    *,
    read_contents: bool = True,
) -> list[ScannedFile]:
    # 표준 경로 템플릿에 걸리는 파일만 우선 추리고, 필요한 경우에만 표지를 읽는다.
    scanned_files = scan_folder(folder_dir, read_contents=False, folder_policy=folder_policy)
    if not read_contents:
        return scanned_files

    matched_paths = [
        file
        for file in scanned_files
        if find_templates_for_file_path(file.path, folder_dir, path_templates, folder_policy)
    ]
    files_to_read = matched_paths or scanned_files

    files: list[ScannedFile] = []
    for file in files_to_read:
        identity = read_file_identity(file.path) if read_contents else file.identity
        files.append(ScannedFile(file.path, identity))
    return files
