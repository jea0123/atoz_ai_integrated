# 프로젝트명 후보를 문서 형식과 무관하게 비교하기 위한 공통 규칙입니다.
from __future__ import annotations

from difflib import SequenceMatcher
from html import unescape
from pathlib import Path
import re
import unicodedata


PROJECT_TITLE_LABELS = {"사업명", "프로젝트명", "프로젝트 제목"}
PROJECT_TITLE_MATCH_MIN_RATIO = 0.58


def clean_project_title_text(value: str | None) -> str:
    text = unicodedata.normalize("NFKC", unescape(str(value or "")))
    text = re.sub(r"\s+", " ", text)
    return text.strip(" -\t\r\n\"'`<>")


def normalize_project_title_label(value: str | None) -> str:
    return re.sub(r"\s+", "", clean_project_title_text(value))


def is_project_title_label_text(value: str | None) -> bool:
    label = normalize_project_title_label(value)
    return label in {normalize_project_title_label(item) for item in PROJECT_TITLE_LABELS}


def strip_project_title_label(value: str | None) -> str:
    text = clean_project_title_text(value)
    text = re.sub(r"^(?:사업명|프로젝트\s*명|프로젝트\s*제목)\s*[:：]?\s*", "", text)
    return clean_project_title_text(text)


def project_title_match_key(value: str | None) -> str:
    text = strip_project_title_label(value)
    text = Path(text).stem if "." in text else text
    text = re.sub(r"\b20\d{2}\s*년도?", "", text)
    text = re.sub(r"\b20\d{2}\s*년", "", text)
    text = re.sub(r"[\s_\-.\[\]{}()（）/\\]+", "", text)
    text = re.sub(r"[^0-9A-Za-z가-힣]+", "", text)
    return text.casefold()


def project_title_similarity(candidate: str | None, expected_project_title: str | None) -> float:
    candidate_key = project_title_match_key(candidate)
    expected_key = project_title_match_key(expected_project_title)
    if not candidate_key or not expected_key:
        return 0.0
    if candidate_key == expected_key:
        return 1.0
    if len(candidate_key) >= 6 and len(expected_key) >= 6:
        if candidate_key in expected_key or expected_key in candidate_key:
            return 0.9
    return SequenceMatcher(None, candidate_key, expected_key).ratio()


def project_title_matches_expected(candidate: str | None, expected_project_title: str | None) -> bool:
    return project_title_similarity(candidate, expected_project_title) >= PROJECT_TITLE_MATCH_MIN_RATIO


def best_matching_project_title(
    candidates: list[str] | tuple[str, ...],
    expected_project_title: str | None,
) -> str:
    if not expected_project_title:
        return ""

    best_score = 0.0
    best_value = ""
    seen: set[str] = set()
    for candidate in candidates:
        value = strip_project_title_label(candidate)
        key = project_title_match_key(value)
        if not value or not key or key in seen:
            continue
        seen.add(key)
        score = project_title_similarity(value, expected_project_title)
        if score > best_score:
            best_score = score
            best_value = value

    return best_value if best_score >= PROJECT_TITLE_MATCH_MIN_RATIO else ""

