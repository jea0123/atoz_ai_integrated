# 문서관리표준, 실제 폴더, 매칭 전략을 조립해 최종 매핑 결과를 만듭니다.
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from app_runtime import (
    log_event,
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
    match_files_by_rule_with_ai_fallback,
)
from output_file_check.folder_policy import FolderPolicy
from output_file_check.folder_scanner import scan_folder
from output_file_check.matcher import DEFAULT_MATCH_THRESHOLD
from output_file_check.models import OutputMatch, PathTemplate, ScannedFile, StandardOutput
from output_file_check.normalization import normalize_for_match
from output_file_check.path_template_reader import read_path_templates
from output_file_check.standard_reader import extract_standard_text, normalize_artifact_category, read_standard_outputs


@dataclass(frozen=True)
class FolderMappingResult:
    standard_project_title: str
    artifact_category: str
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
    standard_text = extract_standard_text(standard_file)
    if not standard_text.strip():
        raise RuntimeError(f"문서 텍스트를 추출하지 못했습니다: {standard_file}")
    standard_project_title = project_title_override or read_standard_project_title(standard_file, standard_text)
    threshold = float(fields.get("threshold") or DEFAULT_MATCH_THRESHOLD)
    artifact_category = normalize_artifact_category(fields.get("artifact_category"))
    standard_outputs = read_standard_outputs(standard_file, standard_text, category=artifact_category)
    path_templates = read_path_templates(standard_file, standard_outputs, standard_text)
    outputs = merge_reference_outputs(
        build_outputs_from_path_templates(standard_outputs, path_templates),
        standard_outputs,
    )
    if artifact_category == "management" and not outputs:
        outputs = standard_outputs
    log_standard_extraction(
        artifact_category,
        standard_text,
        standard_outputs,
        path_templates,
        outputs,
    )
    runtime_env = read_runtime_env()
    ollama_url = resolve_runtime_ollama_generate_url(runtime_env)
    model = resolve_runtime_model(runtime_env)

    match_strategy, match_mode = selected_match_mode(fields, ollama_url)
    files = scan_template_files(
        folder_dir,
        path_templates,
        folder_policy,
        read_contents=match_strategy != "ai_first",
        include_unmatched_paths=match_strategy == "rule_ai_fallback",
    )

    matches = match_files(
        outputs,
        files,
        path_templates,
        folder_dir,
        match_strategy=match_strategy,
        threshold=threshold,
        expected_project_title=standard_project_title,
        folder_policy=folder_policy,
        ollama_url=ollama_url,
        model=model,
    )

    return FolderMappingResult(
        standard_project_title=standard_project_title,
        artifact_category=artifact_category,
        outputs=outputs,
        path_templates=path_templates,
        files=files,
        matches=matches,
        match_mode=match_mode,
    )


def merge_reference_outputs(
    template_outputs: list[StandardOutput],
    standard_outputs: list[StandardOutput],
) -> list[StandardOutput]:
    merged: list[StandardOutput] = []
    seen: set[str] = set()
    for output in [*template_outputs, *standard_outputs]:
        key = normalize_for_match(output.output_name)
        if not key or key in seen:
            continue
        merged.append(output)
        seen.add(key)
    return merged


def log_standard_extraction(
    artifact_category: str,
    standard_text: str,
    standard_outputs: list[StandardOutput],
    path_templates: list[PathTemplate],
    outputs: list[StandardOutput],
) -> None:
    id_lines = [
        " ".join(line.split())[:180]
        for line in standard_text.splitlines()
        if "MFDS-" in line
    ][:8]
    log_event(
        "standard.extract",
        artifact_category=artifact_category,
        standard_output_count=len(standard_outputs),
        path_template_count=len(path_templates),
        output_count=len(outputs),
        has_management_marker=bool(re.search(r"관리\s*문서\s*ID", standard_text, re.IGNORECASE)),
        has_output_marker=bool(re.search(r"산출물\s*코드", standard_text, re.IGNORECASE)),
        has_mfds_p=bool(re.search(r"MFDS-P", standard_text, re.IGNORECASE)),
        sample_id_lines=id_lines,
        sample_outputs=[
            {"output_id": output.output_id, "output_name": output.output_name}
            for output in standard_outputs[:8]
        ],
    )


def match_files(
    outputs: list[StandardOutput],
    files: list[ScannedFile],
    path_templates: list[PathTemplate],
    folder_dir: Path,
    *,
    match_strategy: str,
    threshold: float,
    expected_project_title: str,
    folder_policy: FolderPolicy,
    ollama_url: str,
    model: str,
) -> list[OutputMatch]:
    if match_strategy == "ai_first" and ollama_url:
        return match_files_by_ai_first(
            outputs,
            files,
            path_templates,
            folder_dir,
            threshold=threshold,
            expected_project_title=expected_project_title,
            folder_policy=folder_policy,
            ollama_url=ollama_url,
            model=model,
        )
    if match_strategy == "rule_ai_fallback" and ollama_url:
        return match_files_by_rule_with_ai_fallback(
            outputs,
            files,
            path_templates,
            folder_dir,
            threshold=threshold,
            expected_project_title=expected_project_title,
            folder_policy=folder_policy,
            ollama_url=ollama_url,
            model=model,
        )

    return match_files_by_folder_path(
        outputs,
        files,
        path_templates,
        folder_dir,
        threshold=threshold,
        folder_policy=folder_policy,
    )


def scan_template_files(
    folder_dir: Path,
    path_templates: list[PathTemplate],
    folder_policy: FolderPolicy,
    *,
    read_contents: bool = True,
    include_unmatched_paths: bool = False,
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
    read_paths = {file.path for file in files_to_read}

    files: list[ScannedFile] = []
    for file in scanned_files:
        if file.path in read_paths:
            files.append(ScannedFile(file.path, read_file_identity(file.path)))
        elif include_unmatched_paths:
            files.append(file)
    return files
