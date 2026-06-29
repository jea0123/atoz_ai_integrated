# 파일명/표지값 기반 규칙 점수 계산기입니다. AI 우선 모드에서는 최종 대체 결과로 쓰지 않습니다.
from __future__ import annotations

from difflib import SequenceMatcher
import re

from document_update.patterns import OUTPUT_ID_PATTERN_TEXT
from .models import MatchCandidate, ScannedFile, StandardOutput
from .normalization import (
    STANDARD_ID_IN_NAME_PATTERN,
    normalize_for_match,
    output_id_prefix,
    output_name_from_id,
    strip_attachment_tail,
)


DEFAULT_MATCH_THRESHOLD = 0.72
PARENTHETICAL_CONTENT_PATTERN = re.compile(r"\([^()]*\)|\uFF08[^\uFF08\uFF09]*\uFF09")
LEADING_OUTPUT_ID_PATTERN = re.compile(
    rf"^\s*(?:{OUTPUT_ID_PATTERN_TEXT}|\d{{1,6}})[\s_-]*",
    re.IGNORECASE,
)


def score_file(
    output: StandardOutput,
    scanned_file: ScannedFile,
    *,
    use_output_id: bool = False,
) -> MatchCandidate | None:
    file_stem = strip_attachment_tail(scanned_file.stem)
    file_stem_upper = file_stem.upper()
    if not filename_core_matches_output(output, file_stem):
        return None

    file_normalized_values = normalized_filename_variants(file_stem)
    file_normalized = file_normalized_values[0] if file_normalized_values else ""

    if has_document_title_conflict(output, scanned_file) and not filename_matches_output(
        output,
        file_stem,
        use_output_id=use_output_id,
    ):
        return None

    if use_output_id and output.output_id.upper() in file_stem_upper:
        return MatchCandidate(output, scanned_file, 1.0, "파일명 산출물ID 전체 일치")

    prefix = output_id_prefix(output.output_id)
    if use_output_id and prefix and prefix.upper() in file_stem_upper:
        return MatchCandidate(output, scanned_file, 0.98, "파일명 산출물ID prefix 일치")

    conflicting_output_id = use_output_id and has_conflicting_output_id(file_stem, output)
    best_score = 0.0
    best_reason = ""

    if scanned_file.identity:
        content_candidate = score_content(output, scanned_file, use_output_id=use_output_id)
        if content_candidate and content_candidate.score > best_score:
            best_score = content_candidate.score
            best_reason = content_candidate.reason

    for alias in output.aliases or (output.output_name,):
        for alias_normalized in normalized_alias_variants(alias):
            if not alias_normalized:
                continue

            for file_value in file_normalized_values:
                if alias_normalized == file_value:
                    score, reason = 0.97, "산출물명 정확히 일치"
                elif alias_normalized in file_value:
                    score, reason = 0.94, "파일명에 산출물명 포함"
                elif len(file_value) >= 4 and file_value in alias_normalized:
                    score, reason = 0.86, "산출물명에 파일명 포함"
                else:
                    score = 0.0
                    reason = ""

                if score > best_score:
                    best_score = score
                    best_reason = reason

    if not best_reason:
        return None
    if conflicting_output_id and best_score < 0.94:
        return None

    return MatchCandidate(output, scanned_file, best_score, best_reason)


def filename_core_matches_output(output: StandardOutput, file_stem: str) -> bool:
    file_keys = filename_core_keys(file_stem)
    if not file_keys:
        return False
    output_keys = output_core_keys(output)
    return any(
        filename_core_key_matches(file_key, output_key)
        for file_key in file_keys
        for output_key in output_keys
    )


def filename_core_key_matches(file_key: str, output_key: str) -> bool:
    if not file_key or not output_key:
        return False
    if file_key == output_key:
        return True
    if output_key in file_key:
        return True
    return len(file_key) >= 4 and file_key in output_key


def filename_core_keys(value: str) -> set[str]:
    stem = strip_attachment_tail(value)
    stem = strip_leading_output_id(stem)
    return set(normalized_filename_variants(stem))


def output_core_keys(output: StandardOutput) -> set[str]:
    values = [output.output_name, output_name_from_id(output.output_id), *output.aliases]
    keys: set[str] = set()
    for value in values:
        keys.update(normalized_alias_variants(value))
    return {key for key in keys if key}


def strip_leading_output_id(value: str) -> str:
    text = str(value or "").strip(" -_\t\r\n")
    while True:
        updated = LEADING_OUTPUT_ID_PATTERN.sub("", text, count=1).strip(" -_\t\r\n")
        if updated == text:
            return updated
        text = updated


def filename_matches_output(output: StandardOutput, file_stem: str, *, use_output_id: bool = False) -> bool:
    if not filename_core_matches_output(output, file_stem):
        return False

    file_stem_upper = file_stem.upper()
    if use_output_id and output.output_id.upper() in file_stem_upper:
        return True

    prefix = output_id_prefix(output.output_id)
    if use_output_id and prefix and prefix.upper() in file_stem_upper:
        return True

    file_normalized_values = normalized_filename_variants(file_stem)
    return any(
        alias_key and any(alias_key in file_key for file_key in file_normalized_values)
        for alias in output.aliases or (output.output_name,)
        for alias_key in normalized_alias_variants(alias)
    )


def has_conflicting_output_id(file_stem: str, output: StandardOutput) -> bool:
    output_upper = output.output_id.upper()
    output_prefix_upper = output_id_prefix(output.output_id).upper()
    for match in STANDARD_ID_IN_NAME_PATTERN.finditer(file_stem):
        candidate = match.group(0).upper()
        if candidate in output_upper or output_upper in candidate:
            return False
        if output_prefix_upper and candidate.startswith(output_prefix_upper):
            return False
    return bool(STANDARD_ID_IN_NAME_PATTERN.search(file_stem))


def score_content(
    output: StandardOutput,
    scanned_file: ScannedFile,
    *,
    use_output_id: bool = False,
) -> MatchCandidate | None:
    identity = scanned_file.identity
    if identity is None or identity.error:
        return None

    title_normalized = normalize_content_title_for_match(identity.document_title)
    preview_normalized = normalize_content_title_for_match(identity.preview_text)
    output_id_upper = output.output_id.upper()
    preview_upper = identity.preview_text.upper()

    if use_output_id and output_id_upper and output_id_upper in preview_upper:
        return MatchCandidate(output, scanned_file, 1.0, "문서 내부 산출물ID 전체 일치")

    prefix = output_id_prefix(output.output_id)
    if use_output_id and prefix and prefix.upper() in preview_upper:
        return MatchCandidate(output, scanned_file, 0.98, "문서 내부 산출물ID prefix 일치")

    best_score = 0.0
    best_reason = ""
    for alias in output.aliases or (output.output_name,):
        alias_normalized = normalize_content_title_for_match(alias)
        if not alias_normalized:
            continue

        if alias_normalized and alias_normalized == title_normalized:
            score, reason = 1.0, "문서 내부 문서명 정확히 일치"
        elif alias_normalized and alias_normalized in title_normalized:
            score, reason = 0.96, "문서 내부 문서명에 산출물명 포함"
        elif alias_normalized and alias_normalized in preview_normalized:
            score, reason = 0.90, "문서 내부 텍스트에 산출물명 포함"
        else:
            continue

        if score > best_score:
            best_score = score
            best_reason = reason

    if not best_reason:
        return None

    return MatchCandidate(output, scanned_file, best_score, best_reason)


def has_document_title_conflict(output: StandardOutput, scanned_file: ScannedFile) -> bool:
    identity = scanned_file.identity
    if identity is None or identity.error or not identity.document_title:
        return False

    title_normalized = normalize_content_title_for_match(identity.document_title)
    if not title_normalized:
        return False

    aliases = output.aliases or (output.output_name,)
    best_score = 0.0
    for alias in aliases:
        alias_normalized = normalize_content_title_for_match(alias)
        if not alias_normalized:
            continue
        if alias_normalized == title_normalized:
            return False
        if alias_normalized in title_normalized or title_normalized in alias_normalized:
            return False
        best_score = max(best_score, SequenceMatcher(None, alias_normalized, title_normalized).ratio())

    return best_score < 0.72


def normalize_content_title_for_match(value: str) -> str:
    return normalize_for_match(strip_parenthetical_content(value))


def normalized_filename_variants(value: str) -> list[str]:
    variants = [strip_attachment_tail(value)]
    without_parenthetical = strip_parenthetical_content(variants[0])
    if without_parenthetical != variants[0]:
        variants.append(without_parenthetical)
    return unique_normalized_values(variants)


def normalized_alias_variants(value: str) -> list[str]:
    variants = [str(value)]
    without_parenthetical = strip_parenthetical_content(str(value))
    if without_parenthetical != str(value):
        variants.append(without_parenthetical)
    return unique_normalized_values(variants)


def unique_normalized_values(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = normalize_for_match(value)
        if not key or key in seen:
            continue
        result.append(key)
        seen.add(key)
    return result


def strip_parenthetical_content(value: str) -> str:
    text = str(value)
    while True:
        updated = PARENTHETICAL_CONTENT_PATTERN.sub(" ", text)
        if updated == text:
            return updated
        text = updated

