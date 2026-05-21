# 문서관리표준 PDF에서 산출물명과 산출물 ID 목록을 읽습니다.
from __future__ import annotations

from dataclasses import replace
from pathlib import Path
import re

from document_update.patterns import NUMBER_OUTPUT_ID_PATTERN_TEXT, OUTPUT_ID_PATTERN_TEXT, split_output_id_and_name
from document_update.hwpx_text import add_bundled_site_packages, extract_document_text

from .models import StandardOutput
from .normalization import (
    clean_text,
    normalize_for_match,
    output_id_prefix,
    output_name_from_id,
    unique_clean_values,
)


OUTPUT_ID_IN_LINE_PATTERN = re.compile(
    rf"\b(?:{OUTPUT_ID_PATTERN_TEXT}|{NUMBER_OUTPUT_ID_PATTERN_TEXT})(?:-[^\s]+)?(?![A-Za-z0-9])"
)
MANAGEMENT_SECTION_START_PATTERN = re.compile(r"\*\s*관리문서\s*ID")
OUTPUT_SECTION_START_PATTERN = re.compile(r"\*\s*산출물\s*코드")
SECTION_END_PATTERN = re.compile(r"3\.1\.2\s*파일명")
HEADER_OR_FOOTER_PATTERN = re.compile(
    r"수입식품통합정보시스템|에이투지시스템|V\d+\.\d+|\d{4}\.\d{2}\.\d{2}"
)
TABLE_HEADER_PATTERN = re.compile(r"구분|프로세스|산출물명|산출물ID|폴더명|활동명|작업명")
OUTPUT_NAME_ALIASES = {
    "현행아키텍처분석서": ("아키텍처분석서",),
    "총괄시험계획서": ("총괄시험계획",),
    "성능시험계획서": ("성능시험계획", "성능(부하)시험계획"),
}


def extract_pdf_text_with_layout(file_path: Path) -> str:
    # PDF 표 구조가 덜 깨지도록 가능한 경우 pypdf layout 모드로 텍스트를 읽는다.
    """PDF 표 열을 유지하기 위해 가능한 경우 pypdf layout 모드로 읽는다."""
    try:
        from pypdf import PdfReader
    except ImportError:
        add_bundled_site_packages()
        try:
            from pypdf import PdfReader
        except ImportError:
            return extract_document_text(file_path)

    reader = PdfReader(str(file_path))
    pages: list[str] = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text(extraction_mode="layout") or "")
        except TypeError:
            pages.append(page.extract_text() or "")
    return "\n".join(pages)


def extract_standard_text(standard_file: Path) -> str:
    # 표준 파일 형식에 맞는 텍스트 추출기를 선택한다.
    if standard_file.suffix.lower() == ".pdf":
        return extract_pdf_text_with_layout(standard_file)
    return extract_document_text(standard_file)


def extract_output_section(document_text: str) -> str:
    # 표준 전체 텍스트에서 산출물 목록이 있는 구간만 잘라낸다.
    output_starts = [match.start() for match in OUTPUT_SECTION_START_PATTERN.finditer(document_text)]
    management_starts = [match.start() for match in MANAGEMENT_SECTION_START_PATTERN.finditer(document_text)]
    starts = output_starts or management_starts
    if not starts:
        return document_text

    start = min(starts)
    end_match = SECTION_END_PATTERN.search(document_text, start)
    end = end_match.start() if end_match else len(document_text)
    return document_text[start:end]


def read_standard_outputs(standard_file: Path) -> list[StandardOutput]:
    # 문서관리표준에서 산출물 ID/산출물명 목록을 StandardOutput 리스트로 읽는다.
    """문서관리표준에서 관리문서 ID와 산출물 코드 표를 읽는다."""
    text = extract_standard_text(standard_file)
    if not text.strip():
        raise RuntimeError(f"문서 텍스트를 추출하지 못했습니다: {standard_file}")

    section_text = extract_output_section(text)
    outputs = [
        output
        for line in section_text.splitlines()
        for output in [parse_output_line(line)]
        if output is not None
    ]
    return deduplicate_outputs(outputs)


def parse_output_line(raw_line: str) -> StandardOutput | None:
    # 표준의 한 줄에서 산출물 ID, 산출물명, 폴더명을 뽑아낸다.
    line = raw_line.replace("⦁", " ").strip()
    cleaned_line = clean_text(line)
    if not cleaned_line:
        return None

    id_match = OUTPUT_ID_IN_LINE_PATTERN.search(line)
    if not id_match:
        return None

    if is_page_header_or_table_header(cleaned_line):
        return None

    output_id, id_name = split_output_id_and_name(clean_text(id_match.group(0)))
    id_name = clean_text(id_name or output_name_from_id(output_id))
    before = line[: id_match.start()].strip()
    after = line[id_match.end() :].strip()

    before_parts = split_columns(before)
    after_parts = split_columns(after)
    table_output_name = before_parts[-1] if before_parts else ""
    output_name = id_name or table_output_name or output_name_from_id(output_id)

    if not output_name or TABLE_HEADER_PATTERN.fullmatch(output_name):
        output_name = id_name or output_name_from_id(output_id)

    output_name = clean_text(output_name)
    folder_name = clean_text(" ".join(after_parts)) if after_parts else None
    aliases = aliases_for_output_name(output_name, id_name, table_output_name)

    if not output_name and aliases:
        output_name = aliases[0]

    if not output_name:
        return None

    return StandardOutput(
        output_id=output_id,
        output_name=output_name,
        folder_name=folder_name,
        aliases=aliases,
        source_line=cleaned_line,
    )


def aliases_for_output_name(output_name: str, id_name: str = "", table_output_name: str = "") -> tuple[str, ...]:
    # 산출물ID에 문서명이 붙어 있으면 그 값을 기준으로만 매칭한다.
    # 산출물 컬럼 값은 ID에 문서명이 없을 때만 보조명으로 쓴다.
    extra_aliases = OUTPUT_NAME_ALIASES.get(output_name, ())
    if id_name:
        return unique_clean_values((output_name, id_name, *extra_aliases))
    return unique_clean_values((output_name, table_output_name, *extra_aliases))


def split_columns(value: str) -> list[str]:
    # PDF에서 공백 여러 개로 갈라진 표 컬럼을 나눈다.
    return [clean_text(part) for part in re.split(r"\s{2,}", value) if clean_text(part)]


def is_page_header_or_table_header(line: str) -> bool:
    # 페이지 머리말/표 헤더처럼 산출물 행이 아닌 줄을 걸러낸다.
    if HEADER_OR_FOOTER_PATTERN.search(line) and not TABLE_HEADER_PATTERN.search(line):
        return True
    return bool(TABLE_HEADER_PATTERN.search(line) and "MFDS-" not in line)


def deduplicate_outputs(outputs: list[StandardOutput]) -> list[StandardOutput]:
    # PDF 파싱 중 중복으로 잡힌 산출물을 ID prefix와 이름 기준으로 합친다.
    by_key: dict[tuple[str, str], StandardOutput] = {}

    for output in outputs:
        aliases = unique_clean_values(
            (
                *(output.aliases or ()),
                *aliases_for_output_name(output.output_name, output_name_from_id(output.output_id)),
            )
        )
        name_key = normalize_for_match(aliases[0] if aliases else output.output_name)
        key = (output_id_prefix(output.output_id), name_key)
        current = by_key.get(key)
        if current is None:
            by_key[key] = replace(output, aliases=aliases)
            continue

        merged_aliases = unique_clean_values(current.aliases + aliases)
        preferred = output if len(output.output_id) > len(current.output_id) else current
        by_key[key] = replace(
            preferred,
            folder_name=current.folder_name or output.folder_name,
            aliases=merged_aliases,
            source_line=current.source_line or output.source_line,
        )

    return list(by_key.values())
