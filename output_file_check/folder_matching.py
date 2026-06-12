# 산출물 매칭의 핵심입니다. 기본은 규칙 우선이며 필요한 경우에만 AI로 보강합니다.
from __future__ import annotations

from pathlib import Path
import re

from app_runtime import log_event, parse_json_object
from document_update.ollama_client import generate
from output_file_check.content_identity import extract_file_cover_text, find_document_number, read_file_identity
from output_file_check.folder_policy import FolderPolicy, relative_parent_parts
from output_file_check.matcher import score_file, strip_parenthetical_content
from output_file_check.models import FileIdentity, MatchCandidate, OutputMatch, PathTemplate, ScannedFile, StandardOutput
from output_file_check.normalization import normalize_for_match, strip_attachment_tail


AI_PREVIEW_CHARS = 900
AI_LOG_PREVIEW_CHARS = 700
FOLDER_AI_TIMEOUT_SECONDS = 8


def build_outputs_from_path_templates(
    standard_outputs: list[StandardOutput],
    path_templates: list[PathTemplate],
) -> list[StandardOutput]:
    """표준 산출물 중 폴더 경로 표에 실제로 등장한 산출물만 골라 매칭 대상으로 만든다."""
    output_by_name = index_outputs_by_name(standard_outputs)

    selected: list[StandardOutput] = []
    seen: set[str] = set()
    for template in path_templates:
        key = normalize_for_check_key(template.output_name)
        if not key:
            continue
        output = output_by_name.get(key)
        if output is None:
            continue
        output_key = normalize_for_check_key(output.output_name)
        if output_key in seen:
            continue
        selected.append(output)
        seen.add(output_key)

    return selected


def index_outputs_by_name(outputs: list[StandardOutput]) -> dict[str, StandardOutput]:
    """산출물명과 별칭으로 StandardOutput을 빠르게 찾기 위한 조회 테이블을 만든다."""
    output_by_name: dict[str, StandardOutput] = {}
    for output in outputs:
        for name in (output.output_name, *output.aliases):
            key = normalize_for_check_key(name)
            if key and key not in output_by_name:
                output_by_name[key] = output
    return output_by_name


def match_files_by_folder_path(
    outputs: list[StandardOutput],
    files: list[ScannedFile],
    path_templates: list[PathTemplate],
    root_folder: Path,
    *,
    threshold: float,
    folder_policy: FolderPolicy | None,
) -> list[OutputMatch]:
    """사용자가 규칙 기반을 선택했을 때만 쓰는 경로/파일명 기반 매칭 흐름이다."""
    output_by_name = index_outputs_by_name(outputs)
    templates = [
        template
        for template in path_templates
        if normalize_for_check_key(template.output_name) in output_by_name
    ]
    candidates_by_output: dict[str, list[MatchCandidate]] = {
        normalize_for_check_key(output.output_name): []
        for output in outputs
    }

    for file in files:
        template = find_template_for_file_path(file.path, root_folder, templates, folder_policy)
        if template is None:
            continue

        output = output_by_name.get(normalize_for_check_key(template.output_name))
        if output is None:
            continue

        candidate = score_file(output, file)
        if candidate is None or candidate.score < threshold:
            continue

        candidates_by_output.setdefault(normalize_for_check_key(output.output_name), []).append(candidate)

    return output_matches_from_candidates(outputs, candidates_by_output)


def match_files_by_ai_first(
    outputs: list[StandardOutput],
    files: list[ScannedFile],
    path_templates: list[PathTemplate],
    root_folder: Path,
    *,
    threshold: float,
    expected_project_title: str,
    folder_policy: FolderPolicy | None,
    ollama_url: str,
    model: str,
    all_outputs: list[StandardOutput] | None = None,
) -> list[OutputMatch]:
    """AI 우선 매칭의 메인 흐름이다. 규칙은 AI에 보낼 후보 축소에만 관여한다."""
    output_by_name = index_outputs_by_name(outputs)
    all_outputs_for_filter = all_outputs or outputs
    candidates_by_output = empty_candidate_map(outputs)
    files_by_output = empty_file_scope_map(outputs)
    usable_templates = usable_templates_for_outputs(path_templates, output_by_name)
    for file in files:
        file_for_ai = file_with_cover_text(file)
        if not has_required_cover_text(file_for_ai):
            log_event("ai_match.skipped", file=str(file_for_ai.path), reason="missing_cover_text")
            continue

        title_outputs_all = outputs_matching_document_title(file_for_ai, all_outputs_for_filter)
        title_outputs = [
            output
            for output in title_outputs_all
            if normalize_for_check_key(output.output_name) in output_by_name
        ]
        if title_outputs_all and not title_outputs:
            log_event("ai_match.skipped", file=str(file_for_ai.path), reason="other_output_title")
            continue
        if len(title_outputs) == 1:
            output = title_outputs[0]
            key = normalize_for_check_key(output.output_name)
            candidates_by_output.setdefault(key, []).append(
                MatchCandidate(
                    output,
                    file_with_cover_identity(file_for_ai, output),
                    1.0,
                    "표지 문서명 일치",
                    ai_confidence=1.0,
                )
            )
            continue

        path_outputs = outputs_for_templates(
            find_templates_for_file_path(file_for_ai.path, root_folder, usable_templates, folder_policy),
            output_by_name,
        )
        file_name_outputs = outputs_matching_file_name(file_for_ai, outputs)
        cover_outputs = outputs_matching_cover_text(file_for_ai, outputs)
        candidate_outputs = title_outputs or merge_outputs(file_name_outputs, path_outputs, cover_outputs)
        if not candidate_outputs:
            log_event("ai_match.skipped", file=str(file_for_ai.path), reason="no_output_hint")
            continue

        for output in candidate_outputs:
            key = normalize_for_check_key(output.output_name)
            files_by_output.setdefault(key, {})[file_for_ai.path] = file_for_ai

    selected_ai_paths = {
        file.path
        for output_files in files_by_output.values()
        for file in output_files.values()
    }
    log_event(
        "ai_match.scope",
        files=len(files),
        selected_files=len(selected_ai_paths),
        outputs=len(outputs),
    )

    ai_batch_failed = False
    for output_index, output in enumerate(outputs, start=1):
        if ai_batch_failed:
            break
        key = normalize_for_check_key(output.output_name)
        output_files = list(files_by_output.get(key, {}).values())
        if not output_files:
            continue
        scored_candidates, ai_batch_failed = ai_score_files_for_output(
            output,
            output_files,
            root_folder,
            expected_project_title=expected_project_title,
            ollama_url=ollama_url,
            model=model,
            all_outputs=all_outputs_for_filter,
        )
        candidates_by_output.setdefault(key, []).extend(scored_candidates)
        if ai_batch_failed:
            log_event(
                "ai_match.stopped_after_error",
                output_id=output.output_id,
                output_name=output.output_name,
                remaining_outputs=max(len(outputs) - output_index, 0),
            )

    return output_matches_from_candidates(outputs, candidates_by_output)


def match_files_by_rule_with_ai_fallback(
    outputs: list[StandardOutput],
    files: list[ScannedFile],
    path_templates: list[PathTemplate],
    root_folder: Path,
    *,
    threshold: float,
    expected_project_title: str,
    folder_policy: FolderPolicy | None,
    ollama_url: str,
    model: str,
) -> list[OutputMatch]:
    """규칙 매칭을 먼저 실행하고, 매칭이 비어 있는 산출물만 AI로 보강한다."""
    rule_matches = match_files_by_folder_path(
        outputs,
        files,
        path_templates,
        root_folder,
        threshold=threshold,
        folder_policy=folder_policy,
    )
    missing_outputs = [match.output for match in rule_matches if not match.candidates]
    if not missing_outputs or not ollama_url:
        return rule_matches

    claimed_paths = {
        candidate.file.path
        for match in rule_matches
        for candidate in match.candidates
    }
    fallback_files = [file for file in files if file.path not in claimed_paths]
    if not fallback_files:
        return rule_matches

    log_event(
        "ai_match.fallback_scope",
        outputs=len(outputs),
        missing_outputs=len(missing_outputs),
        files=len(files),
        fallback_files=len(fallback_files),
    )
    ai_matches = match_files_by_ai_first(
        missing_outputs,
        fallback_files,
        path_templates,
        root_folder,
        threshold=threshold,
        expected_project_title=expected_project_title,
        folder_policy=folder_policy,
        ollama_url=ollama_url,
        model=model,
        all_outputs=outputs,
    )
    return merge_rule_matches_with_ai_fallback(rule_matches, ai_matches)


def merge_rule_matches_with_ai_fallback(
    rule_matches: list[OutputMatch],
    ai_matches: list[OutputMatch],
) -> list[OutputMatch]:
    ai_by_output = {
        normalize_for_check_key(match.output.output_name): match
        for match in ai_matches
        if match.candidates
    }
    merged: list[OutputMatch] = []
    for match in rule_matches:
        if match.candidates:
            merged.append(match)
            continue

        fallback = ai_by_output.get(normalize_for_check_key(match.output.output_name))
        merged.append(fallback or match)
    return merged


def empty_candidate_map(outputs: list[StandardOutput]) -> dict[str, list[MatchCandidate]]:
    return {
        normalize_for_check_key(output.output_name): []
        for output in outputs
    }


def empty_file_scope_map(outputs: list[StandardOutput]) -> dict[str, dict[Path, ScannedFile]]:
    return {
        normalize_for_check_key(output.output_name): {}
        for output in outputs
    }


def usable_templates_for_outputs(
    path_templates: list[PathTemplate],
    output_by_name: dict[str, StandardOutput],
) -> list[PathTemplate]:
    return [
        template
        for template in path_templates
        if normalize_for_check_key(template.output_name) in output_by_name
    ]


def output_matches_from_candidates(
    outputs: list[StandardOutput],
    candidates_by_output: dict[str, list[MatchCandidate]],
) -> list[OutputMatch]:
    """산출물별 후보 목록을 점수순으로 정렬해서 화면/CLI용 결과 객체로 묶는다."""
    candidates_by_output = filter_and_dedupe_candidates(outputs, candidates_by_output)
    results: list[OutputMatch] = []
    for output in outputs:
        key = normalize_for_check_key(output.output_name)
        candidates = tuple(sorted(candidates_by_output.get(key, []), key=lambda item: item.score, reverse=True))
        results.append(OutputMatch(output, candidates))
    return results


def filter_and_dedupe_candidates(
    outputs: list[StandardOutput],
    candidates_by_output: dict[str, list[MatchCandidate]],
) -> dict[str, list[MatchCandidate]]:
    filtered: dict[str, list[MatchCandidate]] = {
        normalize_for_check_key(output.output_name): []
        for output in outputs
    }
    selected_by_path: dict[Path, tuple[str, MatchCandidate, tuple[int, float, int]]] = {}

    for output in outputs:
        key = normalize_for_check_key(output.output_name)
        for candidate in candidates_by_output.get(key, []):
            rank = candidate_output_rank(candidate, output, outputs)
            if rank[0] <= 0:
                log_event(
                    "ai_match.candidate_filtered",
                    file=str(candidate.file.path),
                    output_id=output.output_id,
                    output_name=output.output_name,
                    reason="other_output_hint",
                )
                continue

            current = selected_by_path.get(candidate.file.path)
            if current is None or rank > current[2]:
                if current is not None:
                    log_event(
                        "ai_match.candidate_deduped",
                        file=str(candidate.file.path),
                        removed_output=current[0],
                        kept_output=key,
                    )
                selected_by_path[candidate.file.path] = (key, candidate, rank)
            else:
                log_event(
                    "ai_match.candidate_deduped",
                    file=str(candidate.file.path),
                    removed_output=key,
                    kept_output=current[0],
                )

    for key, candidate, _rank in selected_by_path.values():
        filtered.setdefault(key, []).append(candidate)
    return filtered


def candidate_output_rank(
    candidate: MatchCandidate,
    output: StandardOutput,
    outputs: list[StandardOutput],
) -> tuple[int, float, int]:
    output_key = normalize_for_check_key(output.output_name)
    title_outputs = outputs_matching_document_title(candidate.file, outputs)
    if title_outputs:
        title_keys = {normalize_for_check_key(item.output_name) for item in title_outputs}
        if output_key not in title_keys:
            return (0, candidate.score, len(output.output_name))
        return (4, candidate.score, len(output.output_name))

    file_name_outputs = outputs_matching_file_name(candidate.file, outputs)
    if file_name_outputs:
        file_name_keys = {normalize_for_check_key(item.output_name) for item in file_name_outputs}
        if output_key not in file_name_keys:
            return (0, candidate.score, len(output.output_name))
        return (3, candidate.score, len(output.output_name))

    cover_outputs = outputs_matching_cover_text(candidate.file, outputs)
    if cover_outputs:
        cover_keys = {normalize_for_check_key(item.output_name) for item in cover_outputs}
        if output_key in cover_keys:
            return (2, candidate.score, len(output.output_name))

    return (1, candidate.score, len(output.output_name))


def file_with_cover_text(file: ScannedFile) -> ScannedFile:
    """AI 호출 전에 파일 첫 장/표지 텍스트와 표지 문서명을 준비한다."""
    if file.identity and file.identity.preview_text.strip():
        return file
    try:
        identity = read_file_identity(file.path)
        cover_text = identity.preview_text or extract_file_cover_text(file.path)
    except Exception as exc:
        return ScannedFile(file.path, FileIdentity(error=str(exc)))
    log_event(
        "ai_match.cover_text",
        file=str(file.path),
        chars=len(cover_text),
        document_title=identity.document_title,
        preview=log_preview(cover_text),
    )
    return ScannedFile(
        file.path,
        FileIdentity(
            project_title=identity.project_title,
            document_title=identity.document_title,
            document_number=identity.document_number or find_document_number(cover_text),
            preview_text=cover_text,
            error=identity.error,
        ),
    )


def has_required_cover_text(file: ScannedFile) -> bool:
    """AI 모드에서는 첫 장/표지 텍스트가 있는 파일만 판단한다."""
    identity = file.identity
    return bool(identity and identity.preview_text.strip())


def merge_outputs(*groups: list[StandardOutput]) -> list[StandardOutput]:
    merged: list[StandardOutput] = []
    seen: set[str] = set()
    for group in groups:
        for output in group:
            key = normalize_for_check_key(output.output_name)
            if not key or key in seen:
                continue
            merged.append(output)
            seen.add(key)
    return merged


def outputs_matching_file_name(file: ScannedFile, outputs: list[StandardOutput]) -> list[StandardOutput]:
    return outputs_matching_text(file.path.name, outputs)


def outputs_matching_document_title(file: ScannedFile, outputs: list[StandardOutput]) -> list[StandardOutput]:
    identity = file.identity
    if identity is None or not identity.document_title:
        return []
    return outputs_matching_text(identity.document_title, outputs, strip_parenthetical=True)


def outputs_matching_cover_text(file: ScannedFile, outputs: list[StandardOutput]) -> list[StandardOutput]:
    identity = file.identity
    if identity is None:
        return []
    return outputs_matching_text(identity.preview_text, outputs, strip_parenthetical=True)


def outputs_matching_text(text: str, outputs: list[StandardOutput], *, strip_parenthetical: bool = False) -> list[StandardOutput]:
    text_keys = normalize_text_variants_for_match(text, strip_parenthetical=strip_parenthetical)
    if not text_keys:
        return []
    matched: list[StandardOutput] = []
    for output in outputs:
        if any(
            text_matches_alias_key(text_key, alias_key)
            for text_key in text_keys
            for alias_key in output_match_alias_keys(output, strip_parenthetical=strip_parenthetical)
        ):
            matched.append(output)
    return matched


def output_match_alias_keys(output: StandardOutput, *, strip_parenthetical: bool = False) -> list[str]:
    canonical_key = normalize_text_for_match(output.output_name, strip_parenthetical=strip_parenthetical)
    keys: list[str] = []
    seen: set[str] = set()
    for alias in (output.output_name, *output.aliases):
        for alias_key in normalize_text_variants_for_match(alias, strip_parenthetical=strip_parenthetical):
            if not alias_key or alias_key in seen:
                continue
            if alias_key != canonical_key and is_too_generic_alias(alias_key, canonical_key):
                continue
            keys.append(alias_key)
            seen.add(alias_key)
    return keys


def normalize_text_variants_for_match(value: str, *, strip_parenthetical: bool = False) -> list[str]:
    variants = [value]
    without_attachment_tail = strip_attachment_tail(value)
    if without_attachment_tail and without_attachment_tail != value:
        variants.append(without_attachment_tail)

    without_parenthetical = remove_parenthetical_text(value) if strip_parenthetical else value
    if strip_parenthetical and without_parenthetical != value:
        variants.append(without_parenthetical)

    keys: list[str] = []
    seen: set[str] = set()
    for variant in variants:
        key = normalize_for_match(variant)
        if key and key not in seen:
            keys.append(key)
            seen.add(key)
    return keys


def normalize_text_for_match(value: str, *, strip_parenthetical: bool = False) -> str:
    text = remove_parenthetical_text(value) if strip_parenthetical else value
    return normalize_for_match(text)


def remove_parenthetical_text(value: str) -> str:
    return strip_parenthetical_content(value)


def is_too_generic_alias(alias_key: str, canonical_key: str) -> bool:
    if not canonical_key:
        return True
    if len(alias_key) < max(5, len(canonical_key) - 1):
        return True
    generic_terms = {"시험결과서", "결과서", "보고서", "계획서", "매뉴얼", "시나리오"}
    return alias_key in generic_terms


def outputs_for_templates(
    path_templates: list[PathTemplate],
    output_by_name: dict[str, StandardOutput],
) -> list[StandardOutput]:
    """파일 경로에 걸린 표준 경로 템플릿을 산출물 후보 목록으로 바꾼다."""
    outputs: list[StandardOutput] = []
    seen: set[str] = set()
    for template in path_templates:
        output = output_by_name.get(normalize_for_check_key(template.output_name))
        if output is None:
            continue
        key = normalize_for_check_key(output.output_name)
        if key in seen:
            continue
        outputs.append(output)
        seen.add(key)
    return outputs


def ai_score_files_for_output(
    output: StandardOutput,
    files: list[ScannedFile],
    root_folder: Path,
    *,
    expected_project_title: str,
    ollama_url: str,
    model: str,
    all_outputs: list[StandardOutput],
) -> tuple[list[MatchCandidate], bool]:
    """산출물 하나의 후보 파일 묶음을 AI가 한 번에 보고 실제 반영 대상을 고른다."""
    ai_batch_failed = False
    try:
        selected_indexes = ask_ai_for_output_file_matches(
            output,
            files,
            root_folder,
            expected_project_title,
            ollama_url,
            model,
            all_outputs,
        )
    except Exception as exc:
        ai_batch_failed = True
        selected_indexes = repair_under_selected_indexes(output, files, [], all_outputs)
        log_event(
            "ai_match.batch_error",
            output_id=output.output_id,
            output_name=output.output_name,
            error=str(exc),
            recovered_indexes=selected_indexes,
        )

    candidates: list[MatchCandidate] = []
    for index in selected_indexes:
        if index < 1 or index > len(files):
            continue
        file = file_with_cover_identity(files[index - 1], output)
        candidate = MatchCandidate(
            output,
            file,
            1.0,
            "AI 묶음 판단",
            ai_confidence=1.0,
        )
        log_event(
            "ai_match.matched",
            file=str(file.path),
            output_id=output.output_id,
            confidence=1.0,
            score=1.0,
            source="ai_batch",
        )
        candidates.append(candidate)
    return candidates, ai_batch_failed


def file_with_cover_identity(file: ScannedFile, output: StandardOutput) -> ScannedFile:
    """표지 텍스트에서 확정한 산출물명을 화면용 identity에 반영한다."""
    current = file.identity or FileIdentity()
    return ScannedFile(
        file.path,
        FileIdentity(
            project_title=current.project_title,
            document_title=current.document_title or output.output_name,
            document_number=current.document_number,
            preview_text=current.preview_text,
            error=current.error,
        ),
    )

def ask_ai_for_output_file_matches(
    output: StandardOutput,
    files: list[ScannedFile],
    root_folder: Path,
    expected_project_title: str,
    ollama_url: str,
    model: str,
    all_outputs: list[StandardOutput],
) -> list[int]:
    """AI에게 산출물 하나와 후보 파일 묶음을 주고 실제 매칭 파일 index 목록을 받는다."""
    file_lines: list[str] = []
    for index, file in enumerate(files, start=1):
        try:
            relative_path = str(file.path.relative_to(root_folder))
        except ValueError:
            relative_path = str(file.path)
        identity = file.identity or FileIdentity()
        cover_text = " ".join(identity.preview_text.split())[:260]
        file_lines.append(
            f"{index}. file_name={file.path.name}\n"
            f"   path={relative_path}\n"
            f"   cover_text={cover_text}"
        )

    prompt = f"""
You are matching project output files.
Output target:
- output_id: {output.output_id}
- output_name: {output.output_name}
- expected_project_title: {expected_project_title}

Select every matching file index. Do not stop after a few files.
Exclude only files whose cover_text says another document type.
Return only compact JSON with numeric indexes. Never return file names.
Good: {{"selected":[1,2,3,4]}}
Bad: {{"selected":["file.xlsx"]}}

Candidate files:
{chr(10).join(file_lines)}
"""
    log_event(
        "ai_match.batch_request",
        output_id=output.output_id,
        output_name=output.output_name,
        file_count=len(files),
        expected_indexes=list(range(1, len(files) + 1)),
        timeout_seconds=FOLDER_AI_TIMEOUT_SECONDS,
        files=[
            {
                "index": index,
                "path": str(file.path),
                "cover_preview": log_preview((file.identity.preview_text if file.identity else ""), limit=220),
            }
            for index, file in enumerate(files, start=1)
        ],
    )
    raw_text = generate(
        ollama_url,
        model,
        prompt,
        timeout=FOLDER_AI_TIMEOUT_SECONDS,
        options={"temperature": 0, "num_predict": 260},
        response_format="json",
    )
    log_event("ai_match.batch_raw_response", output_id=output.output_id, output_name=output.output_name, raw=log_preview(raw_text))
    parsed = parse_json_object(raw_text)
    selected = parse_selected_indexes(parsed.get("selected") if parsed else None, files)
    if not selected:
        selected = parse_selected_indexes_from_text(raw_text, files)
    selected = repair_under_selected_indexes(output, files, selected, all_outputs)
    selected = filter_selected_indexes_for_output(output, files, selected, all_outputs)
    log_event("ai_match.batch_parsed_response", output_id=output.output_id, output_name=output.output_name, selected=selected)
    return selected


def parse_selected_indexes(value: object, files: list[ScannedFile]) -> list[int]:
    if not isinstance(value, list):
        return []
    indexes: list[int] = []
    for item in value:
        index = parse_selected_item(item, files)
        if index:
            indexes.append(index)
    return sorted(set(indexes))


def repair_under_selected_indexes(
    output: StandardOutput,
    files: list[ScannedFile],
    selected: list[int],
    all_outputs: list[StandardOutput],
) -> list[int]:
    """AI가 일부만 고른 경우, 표지에 산출물명이 명확한 후보는 누락되지 않게 보정한다."""
    cover_indexes = [
        index
        for index, file in enumerate(files, start=1)
        if cover_text_matches_output(file, output, all_outputs)
    ]
    if not cover_indexes:
        return selected
    repaired = sorted(set(selected) | set(cover_indexes))
    if repaired != selected:
        log_event(
            "ai_match.batch_repaired",
            output_id=output.output_id,
            output_name=output.output_name,
            selected_before=selected,
            cover_indexes=cover_indexes,
            selected_after=repaired,
        )
    return repaired


def filter_selected_indexes_for_output(
    output: StandardOutput,
    files: list[ScannedFile],
    selected: list[int],
    all_outputs: list[StandardOutput],
) -> list[int]:
    filtered = [
        index
        for index in selected
        if 1 <= index <= len(files) and cover_text_matches_output(files[index - 1], output, all_outputs)
    ]
    if filtered != selected:
        log_event(
            "ai_match.batch_filtered",
            output_id=output.output_id,
            output_name=output.output_name,
            selected_before=selected,
            selected_after=filtered,
        )
    return filtered


def cover_text_matches_output(
    file: ScannedFile,
    output: StandardOutput,
    all_outputs: list[StandardOutput],
) -> bool:
    identity = file.identity
    if identity is None:
        return False
    output_key = normalize_for_check_key(output.output_name)
    title_outputs = outputs_matching_document_title(file, all_outputs)
    if title_outputs:
        title_keys = {normalize_for_check_key(item.output_name) for item in title_outputs}
        return output_key in title_keys

    if outputs_matching_file_name(file, [output]):
        return True

    cover_keys = normalize_text_variants_for_match(identity.preview_text)
    if not cover_keys:
        return False
    return any(
        text_matches_alias_key(cover_key, alias_key)
        for cover_key in cover_keys
        for alias_key in output_match_alias_keys(output)
    )


def text_matches_alias_key(text_key: str, alias_key: str) -> bool:
    """Match compact titles, with parenthetical variants handled before normalization."""
    if not text_key or not alias_key:
        return False
    if alias_key in text_key:
        return True
    return False


def parse_selected_item(item: object, files: list[ScannedFile]) -> int:
    if isinstance(item, int):
        return item
    if isinstance(item, float) and item.is_integer():
        return int(item)

    text = str(item).strip()
    if re.fullmatch(r"\d+", text):
        return int(text)
    return find_file_index_by_ai_token(text, files)


def parse_selected_indexes_from_text(raw_text: str, files: list[ScannedFile]) -> list[int]:
    match = re.search(r'"selected"\s*:\s*\[([^\]]*)', raw_text)
    if not match:
        return []
    body = match.group(1)
    quoted_tokens = re.findall(r'"([^"]+)"', body)
    if quoted_tokens:
        return sorted({
            index
            for token in quoted_tokens
            for index in [find_file_index_by_ai_token(token, files)]
            if index
        })
    return sorted({int(item) for item in re.findall(r"\b\d+\b", body)})


def find_file_index_by_ai_token(token: str, files: list[ScannedFile]) -> int:
    token_key = normalize_for_match(token)
    if not token_key:
        return 0
    for index, file in enumerate(files, start=1):
        if token == file.path.name or token == str(file.path):
            return index
        name_key = normalize_for_match(file.path.name)
        path_key = normalize_for_match(str(file.path))
        if token_key and (token_key == name_key or token_key in path_key):
            return index
    return 0


def log_preview(value: str, *, limit: int = AI_LOG_PREVIEW_CHARS) -> str:
    """로그가 너무 커지지 않도록 한 줄 미리보기로 줄인다."""
    return " ".join(value.split())[:limit]


def find_template_for_file_path(
    file_path: Path,
    root_folder: Path,
    path_templates: list[PathTemplate],
    folder_policy: FolderPolicy | None,
) -> PathTemplate | None:
    """파일 경로와 가장 길게 맞는 표준 경로 템플릿 하나를 찾는다."""
    original_parts = relative_parent_parts(file_path, root_folder)
    parts = folder_policy.comparable_path_parts(original_parts) if folder_policy else original_parts
    best: PathTemplate | None = None
    for template in path_templates:
        if find_contiguous_subpath(parts, template.template_path) == -1:
            continue
        if best is None or len(template.template_path) > len(best.template_path):
            best = template
    return best


def find_templates_for_file_path(
    file_path: Path,
    root_folder: Path,
    path_templates: list[PathTemplate],
    folder_policy: FolderPolicy | None,
) -> list[PathTemplate]:
    """파일 경로에 맞는 표준 경로 템플릿들을 찾고, 가장 구체적인 것만 남긴다."""
    original_parts = relative_parent_parts(file_path, root_folder)
    parts = folder_policy.comparable_path_parts(original_parts) if folder_policy else original_parts
    matches = [
        template
        for template in path_templates
        if find_contiguous_subpath(parts, template.template_path) != -1
    ]
    if not matches:
        return []

    longest = max(len(template.template_path) for template in matches)
    return [template for template in matches if len(template.template_path) == longest]


def find_contiguous_subpath(actual_path: tuple[str, ...], expected_path: tuple[str, ...]) -> int:
    if not expected_path:
        return -1

    actual_keys = [normalize_for_match(part) for part in actual_path]
    expected_keys = [normalize_for_match(part) for part in expected_path]
    expected_length = len(expected_keys)

    for index in range(0, len(actual_keys) - expected_length + 1):
        if actual_keys[index:index + expected_length] == expected_keys:
            return index
    return -1


def normalize_for_check_key(value: str) -> str:
    """산출물명/별칭을 딕셔너리 key로 쓰기 위해 같은 정규화 규칙을 적용한다."""
    return normalize_for_match(value)
