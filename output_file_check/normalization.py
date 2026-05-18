# 산출물명, 파일명, 텍스트를 비교하기 좋게 정규화하는 함수들입니다.
from __future__ import annotations

from pathlib import Path
import re
import unicodedata


ID_PREFIX_PATTERN = re.compile(
    r"^(?P<prefix>[A-Za-z]{2,10}(?:-[A-Za-z0-9]{1,12})*-\d{2})(?:-(?P<name>.+))?$"
)
STANDARD_ID_IN_NAME_PATTERN = re.compile(
    r"[A-Za-z]{2,10}(?:-[A-Za-z0-9]{1,12}){2,6}(?:-[가-힣A-Za-z0-9()]+)?"
)
VERSION_TOKEN_PATTERN = re.compile(r"[_-]?[vV]\d+(?:\.\d+)*")
REQUEST_ID_PATTERN = re.compile(r"SFR-[A-Za-z0-9-]+", re.IGNORECASE)


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


def output_id_prefix(output_id: str) -> str:
    match = ID_PREFIX_PATTERN.match(output_id)
    return match.group("prefix") if match else output_id


def output_name_from_id(output_id: str) -> str:
    match = ID_PREFIX_PATTERN.match(output_id)
    return clean_text(match.group("name") or "") if match else ""


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
