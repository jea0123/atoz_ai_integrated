# 파일명/표지값 기반 규칙 점수 계산기입니다. AI 우선 모드에서는 최종 대체 결과로 쓰지 않습니다.
from __future__ import annotations

from difflib import SequenceMatcher

from .models import MatchCandidate, ScannedFile, StandardOutput
from .normalization import STANDARD_ID_IN_NAME_PATTERN, normalize_for_match, output_id_prefix


DEFAULT_MATCH_THRESHOLD = 0.72


def score_file(
    output: StandardOutput,
    scanned_file: ScannedFile,
) -> MatchCandidate | None:
    file_stem = scanned_file.stem
    file_stem_upper = file_stem.upper()
    file_normalized = normalize_for_match(file_stem)

    if has_document_title_conflict(output, scanned_file):
        return None

    if output.output_id.upper() in file_stem_upper:
        return MatchCandidate(output, scanned_file, 1.0, "파일명 산출물ID 전체 일치")

    prefix = output_id_prefix(output.output_id)
    if prefix and prefix.upper() in file_stem_upper:
        return MatchCandidate(output, scanned_file, 0.98, "파일명 산출물ID prefix 일치")

    conflicting_output_id = has_conflicting_output_id(file_stem, output)
    best_score = 0.0
    best_reason = ""

    if scanned_file.identity:
        content_candidate = score_content(output, scanned_file)
        if content_candidate and content_candidate.score > best_score:
            best_score = content_candidate.score
            best_reason = content_candidate.reason

    for alias in output.aliases or (output.output_name,):
        alias_normalized = normalize_for_match(alias)
        if not alias_normalized:
            continue

        if alias_normalized == file_normalized:
            score, reason = 0.97, "산출물명 정확히 일치"
        elif alias_normalized in file_normalized:
            score, reason = 0.94, "파일명에 산출물명 포함"
        elif len(file_normalized) >= 4 and file_normalized in alias_normalized:
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


def score_content(output: StandardOutput, scanned_file: ScannedFile) -> MatchCandidate | None:
    identity = scanned_file.identity
    if identity is None or identity.error:
        return None

    title_normalized = normalize_for_match(identity.document_title)
    preview_normalized = normalize_for_match(identity.preview_text)
    output_id_upper = output.output_id.upper()
    preview_upper = identity.preview_text.upper()

    if output_id_upper and output_id_upper in preview_upper:
        return MatchCandidate(output, scanned_file, 1.0, "문서 내부 산출물ID 전체 일치")

    prefix = output_id_prefix(output.output_id)
    if prefix and prefix.upper() in preview_upper:
        return MatchCandidate(output, scanned_file, 0.98, "문서 내부 산출물ID prefix 일치")

    best_score = 0.0
    best_reason = ""
    for alias in output.aliases or (output.output_name,):
        alias_normalized = normalize_for_match(alias)
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

    title_normalized = normalize_for_match(identity.document_title)
    if not title_normalized:
        return False

    aliases = output.aliases or (output.output_name,)
    best_score = 0.0
    for alias in aliases:
        alias_normalized = normalize_for_match(alias)
        if not alias_normalized:
            continue
        if alias_normalized == title_normalized:
            return False
        if alias_normalized in title_normalized or title_normalized in alias_normalized:
            return False
        best_score = max(best_score, SequenceMatcher(None, alias_normalized, title_normalized).ratio())

    return best_score < 0.72

