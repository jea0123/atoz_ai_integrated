# 산출물명, 파일명, 텍스트를 비교하기 좋게 정규화하는 함수들입니다.
from __future__ import annotations

from pathlib import Path
import re
import unicodedata

from document_update.patterns import OUTPUT_ID_PATTERN_TEXT, split_output_id_and_name


STANDARD_ID_IN_NAME_PATTERN = re.compile(
    rf"{OUTPUT_ID_PATTERN_TEXT}(?:-[가-힣A-Za-z0-9()]+)?(?![A-Za-z0-9])"
)
VERSION_TOKEN_PATTERN = re.compile(r"[_-]?[vV]\d+(?:\.\d+)*")
REQUEST_ID_PATTERN = re.compile(r"SFR-[A-Za-z0-9-]+", re.IGNORECASE)
ATTACHMENT_BRACKET_TAIL_PATTERN = re.compile(
    r"[\s_-]*[\[\(（［｛]\s*(?:별첨|첨부)\s*\d*[^)\]\}）］｝]*[\)\]\}）］｝].*$",
    re.IGNORECASE,
)
ATTACHMENT_WORD_TAIL_PATTERN = re.compile(r"[\s_-]*(?:별첨|첨부)\s*\d+.*$", re.IGNORECASE)


def compact_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def clean_text(value: str) -> str:
    text = unicodedata.normalize("NFKC", value)
    text = text.replace("⦁", " ").replace("㈜", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip(" -\t\r\n")


def normalize_for_match(value: str) -> str:
    text = unicodedata.normalize("NFKC", value)
    text = Path(text).stem if "." in text else text
    text = VERSION_TOKEN_PATTERN.sub("", text)
    text = REQUEST_ID_PATTERN.sub("", text)
    text = text.replace("테스트", "시험")
    text = re.sub(r"[\s_\-.\[\]{}()（）/\\]+", "", text)
    text = re.sub(r"[^0-9A-Za-z가-힣]+", "", text)
    return text.casefold()


def strip_attachment_tail(value: str) -> str:
    text = str(value)
    stem = Path(text).stem if "." in text else text
    for pattern in (ATTACHMENT_BRACKET_TAIL_PATTERN, ATTACHMENT_WORD_TAIL_PATTERN):
        trimmed = pattern.sub("", stem).strip(" -_\t\r\n")
        if trimmed != stem:
            return trimmed
    return stem


def output_id_prefix(output_id: str) -> str:
    prefix, _name = split_output_id_and_name(clean_text(output_id))
    return prefix or output_id


def output_name_from_id(output_id: str) -> str:
    _prefix, name = split_output_id_and_name(clean_text(output_id))
    return clean_text(name)


def unique_clean_values(values: list[str] | tuple[str, ...]) -> tuple[str, ...]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = clean_text(value)
        key = normalize_for_match(cleaned)
        if not cleaned or key in seen:
            continue
        result.append(cleaned)
        seen.add(key)
    return tuple(result)


def filesystem_safe_stem(value: str) -> str:
    text = clean_text(value)
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", text)
    text = text.strip(" ._")
    return text or "renamed_output"
