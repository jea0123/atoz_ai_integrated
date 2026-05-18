# 후보 파일 첫 장/표지에서 프로젝트명, 문서명, 짧은 표지 텍스트를 읽습니다.
from __future__ import annotations

from html import unescape
from pathlib import Path
import re

from document_update.hwpx_text import (
    extract_text_from_hwp_cover,
    extract_text_from_hwpx_cover,
    extract_text_from_ooxml_cover,
    extract_text_from_pdf_cover,
    is_hwpx_zip,
    is_zip_container,
)
from document_update.excel_ooxml import EXCEL_DOCUMENT_SUFFIXES, extract_excel_cover_text, find_excel_cover_identity

from .models import FileIdentity
from .normalization import clean_text, compact_space, normalize_for_match
from .standard_reader import extract_standard_text


OOXML_SUFFIXES = {
    ".docx",
    ".docm",
    ".dotx",
    ".dotm",
    ".pptx",
    ".pptm",
    ".potx",
    ".potm",
    ".ppsx",
    ".ppsm",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
}
TEXT_READ_SUFFIXES = OOXML_SUFFIXES | {".hwp", ".hwpx", ".pdf"}
COVER_TEXT_CHARS = 5000
TITLE_KEYWORDS = (
    "계획서",
    "결과서",
    "정의서",
    "시나리오",
    "케이스",
    "매뉴얼",
    "보고서",
    "대장",
    "목록",
    "추적표",
    "회의록",
    "WBS",
)
NOISE_PREFIXES = (
    "식품의약품안전처",
    "개 정 이 력",
    "개정 이력",
    "<목 차>",
    "&lt;목 차&gt;",
)
ANGLE_FIELD_PATTERN = re.compile(r"<([^<>]*)>")
METADATA_LABEL_PATTERN = re.compile(
    r"^(?:문서번호|문서버전|개정일자|작성자|버전|개정사유|개정내역|승인)(?:\s|$)"
)


def read_standard_project_title(standard_file: Path) -> str:
    # 문서관리표준에서 프로젝트명/사업명을 추정한다.
    """문서관리표준 표지에서 사업명/프로젝트명을 읽는다."""
    text = extract_standard_text(standard_file)
    first_part = text[:6000]
    lines = meaningful_lines(first_part)

    for index, line in enumerate(lines):
        if "문서관리표준" not in line:
            continue
        before, _, after = line.partition("문서관리표준")
        if before.strip():
            return clean_project_title(before)
        if index > 0:
            return clean_project_title(lines[index - 1])
        if after.strip():
            return clean_project_title(after)

    for line in lines:
        match = re.search(r"(?:사업명|프로젝트\s*명)\s*[:：]?\s*(.+)", line)
        if match:
            return clean_project_title(match.group(1))

    return ""


def read_file_identity(file_path: Path) -> FileIdentity:
    # 후보 파일 첫 장/표지에서 프로젝트명, 문서명, 짧은 표지 텍스트를 만든다.
    """파일 내부 텍스트에서 프로젝트명/사업명과 문서명을 추정한다."""
    if file_path.suffix.lower() not in TEXT_READ_SUFFIXES:
        return FileIdentity(error=f"지원하지 않는 내부 읽기 형식: {file_path.suffix}")

    cover_project_title = ""
    cover_document_title = ""
    if file_path.suffix.lower() in EXCEL_DOCUMENT_SUFFIXES:
        try:
            cover_project_title, cover_document_title = find_excel_cover_identity(file_path)
        except Exception:
            cover_project_title, cover_document_title = "", ""
        if cover_project_title or cover_document_title:
            return FileIdentity(
                project_title=cover_project_title,
                document_title=cover_document_title,
                preview_text=compact_space(f"{cover_project_title}\n{cover_document_title}"),
            )

    try:
        text = extract_file_cover_text(file_path)
    except Exception as exc:
        return FileIdentity(error=str(exc))

    if not text.strip():
        return FileIdentity(error="파일 내부 텍스트를 읽지 못했습니다.")

    project_title, document_title = parse_identity_from_text(text)
    return FileIdentity(
        project_title=cover_project_title or project_title,
        document_title=cover_document_title or document_title,
        preview_text=compact_space(text[:COVER_TEXT_CHARS]),
    )


def clean_identity_value(value: str) -> str:
    # 표지에서 읽은 프로젝트명/문서명 후보의 공백과 꺾쇠 기호를 정리한다.
    return value.strip().strip("\"'`<>")


def meaningful_cover_lines(document_text: str) -> list[str]:
    # 표지 앞부분에서 프로젝트명/문서명으로 쓸 만한 줄만 남긴다.
    noise_prefixes = (
        "식품의약품안전처",
        "㈜",
        "개 정 이 력",
        "<목 차>",
        "&lt;목 차&gt;",
    )

    lines: list[str] = []
    for line in document_text.splitlines():
        value = clean_identity_value(line)
        if not value:
            continue
        if any(value.startswith(prefix) for prefix in noise_prefixes):
            continue
        lines.append(value)

    return lines


def find_target_identity_by_rule(document_text: str) -> tuple[str, str]:
    # 표지의 일반적인 프로젝트명, 문서명, 문서번호 순서를 기준으로 읽는다.
    lines = meaningful_cover_lines(document_text)

    for index, line in enumerate(lines):
        if line == "문서번호" and index >= 2:
            return lines[index - 2], lines[index - 1]

    if len(lines) >= 2:
        return lines[0], lines[1]

    return "", ""


def extract_file_cover_text(file_path: Path) -> str:
    # 파일 형식별로 전체 본문 대신 첫 장/표지에 가까운 텍스트만 읽는다.
    suffix = file_path.suffix.lower()

    if suffix == ".pdf":
        return extract_text_from_pdf_cover(file_path, max_chars=COVER_TEXT_CHARS)

    if is_hwpx_zip(file_path):
        return extract_text_from_hwpx_cover(file_path, max_chars=COVER_TEXT_CHARS)

    if suffix in EXCEL_DOCUMENT_SUFFIXES:
        return extract_excel_cover_text(file_path, max_chars=1000)

    if suffix in {".hwp", ".hwpx"}:
        return extract_text_from_hwp_cover(file_path, max_chars=COVER_TEXT_CHARS)

    if suffix in OOXML_SUFFIXES:
        if not is_zip_container(file_path):
            raise RuntimeError(f"Office XML 형식이 아닙니다: {file_path.name}")
        return extract_text_from_ooxml_cover(file_path, max_chars=COVER_TEXT_CHARS)

    raise RuntimeError(f"지원하지 않는 내부 읽기 형식: {file_path.suffix}")


def parse_identity_from_text(document_text: str) -> tuple[str, str]:
    # 표지 텍스트에서 프로젝트명과 문서명을 여러 힌트 순서로 추정한다.
    """표지형 문서와 PPT/매뉴얼류의 첫 텍스트를 함께 처리한다."""
    rule_project_title, rule_document_title = find_target_identity_by_rule(document_text)
    lines = meaningful_lines(document_text[:8000])
    cover_project_title, cover_document_title = find_cover_identity(lines)

    project_title = (
        find_labeled_project_title(lines)
        or cover_project_title
        or clean_project_title_if_valid(rule_project_title)
    )
    document_title = (
        find_labeled_document_title(lines)
        or cover_document_title
        or clean_document_title_if_valid(rule_document_title)
    )

    if not document_title:
        document_title = find_keyword_title(lines)

    if not project_title:
        project_title = find_project_title_near_document_title(lines, document_title)

    return project_title, document_title


def meaningful_lines(document_text: str) -> list[str]:
    # 빈 줄, 목차 잡음, 바이너리 잡음을 빼고 의미 있는 줄만 남긴다.
    lines: list[str] = []
    for raw_line in document_text.splitlines():
        line = normalize_extracted_line(raw_line)
        if not line:
            continue
        if any(line.startswith(prefix) for prefix in NOISE_PREFIXES):
            continue
        if re.fullmatch(r"[\d.\-_/()'\"]+", line):
            continue
        if is_binary_noise_line(line):
            continue
        lines.append(line)
    return lines


def normalize_extracted_line(raw_line: str) -> str:
    # XML/HTML 엔티티와 꺾쇠 필드를 정리해 비교 가능한 한 줄로 만든다.
    text = unescape(raw_line)
    fields = [clean_text(field) for field in ANGLE_FIELD_PATTERN.findall(text)]
    fields = [field for field in fields if field]
    if fields:
        text = " ".join(fields)
    return clean_text(text).strip("<>")


def is_binary_noise_line(line: str) -> bool:
    # 문서 파서가 뱉는 읽기 어려운 깨진 줄인지 판단한다.
    if not line:
        return True
    if re.fullmatch(r"[<>]+", line):
        return True
    if re.fullmatch(r"[\u4e00-\u9fff]{1,12}", line):
        return True
    readable_count = len(re.findall(r"[가-힣A-Za-z0-9]", line))
    if readable_count == 0:
        return True
    return False


def find_cover_identity(lines: list[str]) -> tuple[str, str]:
    # 표지 상단 구조에서 프로젝트명과 문서명을 우선적으로 찾는다.
    """표지 상단의 프로젝트명/문서명을 우선 읽는다."""
    for index, line in enumerate(lines):
        if not line.startswith("문서번호"):
            continue

        candidates = [
            value
            for value in lines[:index]
            if is_cover_value(value)
        ]
        if len(candidates) >= 2:
            return clean_project_title(candidates[-2]), clean_document_title(candidates[-1])

    candidates = [value for value in lines[:8] if is_cover_value(value)]
    if len(candidates) >= 2:
        return clean_project_title(candidates[0]), clean_document_title(candidates[1])

    return "", ""


def is_cover_value(line: str) -> bool:
    # 표지의 실제 값으로 쓸 수 있는 줄인지 필터링한다.
    if not line or len(line) > 120:
        return False
    if any(line.startswith(prefix) for prefix in NOISE_PREFIXES):
        return False
    if METADATA_LABEL_PATTERN.search(line):
        return False
    if line.startswith("(주)") or line.startswith("㈜"):
        return False
    if is_binary_noise_line(line):
        return False
    return True


def find_labeled_project_title(lines: list[str]) -> str:
    # '사업명/프로젝트명' 같은 라벨 주변에서 프로젝트명을 찾는다.
    for index, line in enumerate(lines):
        match = re.search(r"(?:사업명|프로젝트\s*명|프로젝트\s*제목)\s*[:：]\s*(.+)", line)
        if match:
            return clean_project_title(match.group(1))

        if re.fullmatch(r"사업명|프로젝트\s*명|프로젝트\s*제목", line) and index + 1 < len(lines):
            return clean_project_title(lines[index + 1])

    return ""


def find_labeled_document_title(lines: list[str]) -> str:
    # '문서명/문서 제목' 같은 라벨 주변에서 문서명을 찾는다.
    for index, line in enumerate(lines):
        match = re.search(r"(?:문서명|문서\s*제목|파일\s*제목)\s*[:：]\s*(.+)", line)
        if match:
            return clean_document_title(match.group(1))

        if re.fullmatch(r"문서명|문서\s*제목|파일\s*제목", line) and index + 1 < len(lines):
            return clean_document_title(lines[index + 1])

    return ""


def find_keyword_title(lines: list[str]) -> str:
    # 계획서/결과서/매뉴얼 같은 키워드가 있는 짧은 줄을 문서명 후보로 잡는다.
    for line in lines[:40]:
        if len(line) > 120:
            continue
        if not is_cover_value(line):
            continue
        if any(keyword in line for keyword in TITLE_KEYWORDS):
            return clean_document_title(line)
    return ""


def find_project_title_near_document_title(lines: list[str], document_title: str) -> str:
    # 문서명 바로 위쪽 줄에서 프로젝트명처럼 보이는 값을 찾는다.
    if not document_title:
        return ""

    document_key = normalize_for_match(document_title)
    for index, line in enumerate(lines):
        if normalize_for_match(line) != document_key:
            continue
        for candidate in reversed(lines[max(0, index - 4):index]):
            if any(keyword in candidate for keyword in TITLE_KEYWORDS):
                continue
            if len(candidate) <= 120:
                return clean_project_title(candidate)
    return ""


def clean_project_title(value: str) -> str:
    # 프로젝트명 라벨과 불필요한 공백/기호를 제거한다.
    text = clean_text(value)
    text = re.sub(r"^(?:사업명|프로젝트\s*명|프로젝트\s*제목)\s*[:：]?\s*", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" -")


def clean_document_title(value: str) -> str:
    # 문서명 라벨과 불필요한 공백/기호를 제거한다.
    text = clean_text(value)
    text = re.sub(r"^(?:문서명|문서\s*제목|파일\s*제목|제목)\s*[:：]?\s*", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" -")


def clean_project_title_if_valid(value: str) -> str:
    # 정리한 프로젝트명이 표지 값으로 유효할 때만 반환한다.
    text = clean_project_title(value)
    return text if is_cover_value(text) else ""


def clean_document_title_if_valid(value: str) -> str:
    # 정리한 문서명이 표지 값으로 유효할 때만 반환한다.
    text = clean_document_title(value)
    return text if is_cover_value(text) else ""
