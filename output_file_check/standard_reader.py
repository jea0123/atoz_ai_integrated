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
MANAGEMENT_SECTION_START_PATTERN = re.compile(r"(?:\*\s*)?관\s*리\s*문\s*서\s*(?:I\s*D|아이\s*디)", re.IGNORECASE)
OUTPUT_SECTION_START_PATTERN = re.compile(r"(?:\*\s*)?산\s*출\s*물\s*코\s*드", re.IGNORECASE)
SECTION_END_PATTERN = re.compile(r"3\.1\.2\s*파일명")
RUN_PATTERN = re.compile(r"\S(?:.*?\S)?(?=\s{2,}\S|$)")
HEADER_OR_FOOTER_PATTERN = re.compile(
    r"수입식품통합정보시스템|에이투지시스템|V\d+\.\d+|\d{4}\.\d{2}\.\d{2}"
)
TABLE_HEADER_PATTERN = re.compile(r"구분|프로세스|산출물명|산출물ID|폴더명|활동명|작업명")
GENERIC_STAGE_NAMES = {
    "프로젝트",
    "착수",
    "시작",
    "실행",
    "종료",
    "관리",
    "분석",
    "설계",
    "구현",
    "시험",
    "인도",
}
OUTPUT_NAME_ALIASES = {
    "현행아키텍처분석서": ("아키텍처분석서",),
    "총괄시험계획서": ("총괄시험계획",),
    "성능시험계획서": ("성능시험계획", "성능(부하)시험계획"),
}


def normalize_artifact_category(value: str | None) -> str:
    # 표준 추출 범위를 고르는 화면/API 값을 내부 category 값으로 정리한다.
    category = clean_text(value or "").casefold()
    if category in {"management", "manage", "관리", "관리산출물", "관리문서"}:
        return "management"
    if category in {"development", "develop", "dev", "개발", "개발산출물", "산출물"}:
        return "development"
    return "auto"


def extract_pdf_text_with_layout(file_path: Path) -> str:
    # PDF 표 구조가 덜 깨지도록 가능한 경우 pypdf layout 모드로 텍스트를 읽는다.
    """PDF 표 열을 유지하기 위해 가능한 경우 pypdf layout 모드로 읽는다."""
    try:
        import fitz

        with fitz.open(str(file_path)) as document:
            return "\n".join(page.get_text("text") or "" for page in document)
    except Exception:
        pass

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


def extract_output_section(document_text: str, category: str = "auto") -> str:
    # 표준 전체 텍스트에서 산출물 목록이 있는 구간만 잘라낸다.
    category = normalize_artifact_category(category)
    output_matches = list(OUTPUT_SECTION_START_PATTERN.finditer(document_text))
    management_matches = list(MANAGEMENT_SECTION_START_PATTERN.finditer(document_text))
    if category == "management":
        matches = management_matches
    elif category == "development":
        matches = output_matches
    else:
        matches = output_matches or management_matches

    if not matches and category in {"development", "management"}:
        return ""
    if not matches:
        return document_text

    start_match = min(matches, key=lambda match: match.start())
    start = start_match.start()
    end_patterns = (
        (SECTION_END_PATTERN, MANAGEMENT_SECTION_START_PATTERN)
        if category == "management"
        else (SECTION_END_PATTERN, OUTPUT_SECTION_START_PATTERN, MANAGEMENT_SECTION_START_PATTERN)
    )
    end_candidates = [
        match.start()
        for pattern in end_patterns
        for match in pattern.finditer(document_text, start_match.end())
    ]
    end = min(end_candidates) if end_candidates else len(document_text)
    return document_text[start:end]


def read_standard_outputs(
    standard_file: Path,
    standard_text: str | None = None,
    *,
    category: str = "auto",
) -> list[StandardOutput]:
    # 문서관리표준에서 산출물 ID/산출물명 목록을 StandardOutput 리스트로 읽는다.
    """문서관리표준에서 관리문서 ID와 산출물 코드 표를 읽는다."""
    text = standard_text if standard_text is not None else extract_standard_text(standard_file)
    if not text.strip():
        raise RuntimeError(f"문서 텍스트를 추출하지 못했습니다: {standard_file}")

    normalized_category = normalize_artifact_category(category)
    section_text = extract_output_section(text, normalized_category)
    outputs = parse_output_section_lines(
        section_text,
        loose_id=normalized_category in {"management", "development"},
        category=normalized_category,
    )
    outputs = [
        output
        for output in outputs
        if output_matches_category(output, normalized_category)
    ]
    return deduplicate_outputs(outputs)


def output_matches_category(output: StandardOutput, category: str) -> bool:
    return True


def parse_output_section_lines(
    section_text: str,
    *,
    loose_id: bool = False,
    category: str = "auto",
) -> list[StandardOutput]:
    if loose_id:
        sequential_outputs = parse_sequential_output_lines(section_text, category=category)
        if sequential_outputs:
            return sequential_outputs

    columns = find_output_table_columns(section_text) if loose_id else None
    outputs: list[StandardOutput] = []
    for line in section_text.splitlines():
        output = parse_output_line_by_columns(line, columns) if columns else None
        if output is None:
            output = parse_output_line(line, loose_id=loose_id)
        if output is not None:
            outputs.append(output)
    return outputs


def parse_sequential_output_lines(section_text: str, *, category: str) -> list[StandardOutput]:
    lines = [clean_text(line) for line in section_text.splitlines() if clean_text(line)]
    if not lines:
        return []

    lines = drop_pdf_page_header_lines(lines)
    category = normalize_artifact_category(category)
    activity_index = first_normalized_line_index(lines, "활동명")
    if category == "management" and activity_index is not None:
        lines = lines[:activity_index]
    elif category == "development" and activity_index is not None:
        lines = lines[activity_index + 1:]

    id_header_index = first_normalized_line_index(lines, "산출물id")
    if id_header_index is not None:
        lines = lines[id_header_index + 1:]

    outputs: list[StandardOutput] = []
    for index in range(0, len(lines) - 1):
        output_name = clean_text(lines[index])
        output_id_text = clean_text(lines[index + 1])
        if (
            index + 2 < len(lines)
            and not looks_like_sequential_output_id(output_id_text)
            and is_wrapped_output_name_continuation(output_id_text)
            and looks_like_sequential_output_id(lines[index + 2])
        ):
            output_name = clean_text(f"{output_name}{output_id_text}")
            output_id_text = clean_text(lines[index + 2])
        if not is_sequential_output_name(output_name):
            continue
        if not looks_like_sequential_output_id(output_id_text):
            continue

        output_id, id_name = split_output_id_and_name(output_id_text)
        if category == "development" and id_name:
            output_name = clean_text(id_name)
            aliases = aliases_for_output_name(output_name, id_name, lines[index])
        else:
            output_name = clean_text(output_name)
            aliases = aliases_for_output_name(output_name)
        outputs.append(
            StandardOutput(
                output_id=output_id,
                output_name=output_name,
                folder_name=None,
                aliases=aliases,
                source_line=f"{lines[index]} {lines[index + 1]}",
            )
        )
    return outputs


def drop_pdf_page_header_lines(lines: list[str]) -> list[str]:
    result: list[str] = []
    index = 0
    while index < len(lines):
        if is_pdf_page_header_start(lines, index):
            end_index = pdf_page_header_end(lines, index)
            if end_index is not None:
                index = end_index
                continue
        result.append(lines[index])
        index += 1
    return result


def is_pdf_page_header_start(lines: list[str], index: int) -> bool:
    if index + 2 >= len(lines):
        return False
    current = normalize_for_match(lines[index])
    next_line = lines[index + 1]
    following_line = lines[index + 2]
    return (
        current == normalize_for_match("문서관리표준")
        and HEADER_OR_FOOTER_PATTERN.search(next_line) is not None
        and following_line.upper().startswith("MFDS-")
    )


def pdf_page_header_end(lines: list[str], start_index: int) -> int | None:
    for index in range(start_index, min(start_index + 10, len(lines))):
        if "에이투지시스템" in lines[index]:
            return index + 1
    return None


def first_normalized_line_index(lines: list[str], target_key: str) -> int | None:
    for index, line in enumerate(lines):
        if normalize_for_match(line) == target_key:
            return index
    return None


def is_sequential_output_name(value: str) -> bool:
    text = clean_text(value)
    if not text or TABLE_HEADER_PATTERN.search(text):
        return False
    if is_wrapped_output_name_continuation(text):
        return False
    if normalize_for_match(text) in {normalize_for_match(name) for name in GENERIC_STAGE_NAMES}:
        return False
    if looks_like_sequential_output_id(text):
        return False
    return bool(re.search(r"[가-힣]", text) or normalize_for_match(text) in {"wbs"})


def is_wrapped_output_name_continuation(value: str) -> bool:
    text = clean_text(value)
    return bool(re.fullmatch(r"\([^)]*[가-힣][^)]*\)", text))


def looks_like_sequential_output_id(value: str) -> bool:
    text = clean_text(value)
    if not text or TABLE_HEADER_PATTERN.search(text):
        return False
    if len(text) > 100:
        return False
    return bool(re.search(r"\d", text) or re.search(r"ID", text, re.IGNORECASE) or text.upper().startswith("MFDS-"))


def find_output_table_columns(section_text: str) -> dict[str, tuple[int, int | None]] | None:
    for line in section_text.splitlines():
        runs = extract_runs(line)
        if not runs:
            continue
        headers = [(start, normalize_for_match(value)) for start, value in runs]
        name_start = first_header_start(headers, {"산출물명", "문서명"})
        id_start = first_header_start(headers, {"산출물id", "관리문서id", "문서id", "id"})
        if name_start is None or id_start is None:
            continue

        starts = sorted({start for start, _value in headers})
        return {
            "name": (name_start, next_column_start(starts, name_start)),
            "id": (id_start, next_column_start(starts, id_start)),
        }
    return None


def first_header_start(headers: list[tuple[int, str]], candidates: set[str]) -> int | None:
    for start, value in headers:
        if value in candidates:
            return start
    for start, value in headers:
        if any(candidate in value for candidate in candidates):
            return start
    return None


def next_column_start(starts: list[int], current: int) -> int | None:
    for start in starts:
        if start > current:
            return start
    return None


def parse_output_line_by_columns(raw_line: str, columns: dict[str, tuple[int, int | None]] | None) -> StandardOutput | None:
    if not columns:
        return None
    line = raw_line.replace("⦁", " ").rstrip()
    cleaned_line = clean_text(line)
    if not cleaned_line or is_page_header_or_table_header(cleaned_line):
        return None

    output_id = value_in_column(line, *columns["id"])
    output_name = value_in_column(line, *columns["name"])
    if not output_id or not output_name:
        return None
    if not looks_like_loose_output_id(output_id) or not is_loose_output_name(output_name):
        return None

    return StandardOutput(
        output_id=output_id,
        output_name=output_name,
        folder_name=None,
        aliases=aliases_for_output_name(output_name),
        source_line=cleaned_line,
    )


def value_in_column(line: str, start: int, end: int | None) -> str:
    values = [
        value
        for run_start, value in extract_runs(line)
        if run_start >= max(start - 2, 0) and (end is None or run_start < max(end - 2, start))
    ]
    return clean_text(" ".join(values))


def parse_output_line(raw_line: str, *, loose_id: bool = False) -> StandardOutput | None:
    # 표준의 한 줄에서 산출물 ID, 산출물명, 폴더명을 뽑아낸다.
    line = raw_line.replace("⦁", " ").strip()
    cleaned_line = clean_text(line)
    if not cleaned_line:
        return None

    if loose_id:
        loose_output = parse_loose_output_line(line)
        if loose_output is not None:
            return loose_output

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
    trailing_output_name = candidate_output_name_after_id(after_parts)
    output_name = id_name or table_output_name or trailing_output_name or output_name_from_id(output_id)

    if not output_name or TABLE_HEADER_PATTERN.fullmatch(output_name):
        output_name = id_name or trailing_output_name or output_name_from_id(output_id)

    output_name = clean_text(output_name)
    folder_name = clean_text(" ".join(after_parts)) if after_parts else None
    aliases = aliases_for_output_name(output_name, id_name, table_output_name, trailing_output_name)

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


def parse_loose_output_line(raw_line: str) -> StandardOutput | None:
    line = raw_line.replace("⦁", " ").strip()
    cleaned_line = clean_text(line)
    if not cleaned_line or is_page_header_or_table_header(cleaned_line):
        return None

    columns = split_columns(line)
    if len(columns) < 2:
        columns = [clean_text(part) for part in re.split(r"\s+", line) if clean_text(part)]
    if len(columns) < 2:
        return None

    for index, value in enumerate(columns):
        output_id = clean_text(value)
        if not looks_like_loose_output_id(output_id):
            continue

        output_name = next_loose_output_name(columns, index)
        if not output_name:
            continue

        return StandardOutput(
            output_id=output_id,
            output_name=output_name,
            folder_name=None,
            aliases=aliases_for_output_name(output_name),
            source_line=cleaned_line,
        )
    return None


def looks_like_loose_output_id(value: str) -> bool:
    text = clean_text(value)
    if not text or TABLE_HEADER_PATTERN.search(text):
        return False
    if re.fullmatch(r"\d{1,3}", text):
        return False
    if len(text) > 80:
        return False
    if re.search(r"[가-힣]", text):
        return False
    return bool(re.search(r"\d", text))


def next_loose_output_name(columns: list[str], id_index: int) -> str:
    for candidate in columns[id_index + 1:]:
        output_name = clean_text(candidate)
        if is_loose_output_name(output_name):
            return output_name
    for candidate in reversed(columns[:id_index]):
        output_name = clean_text(candidate)
        if is_loose_output_name(output_name):
            return output_name
    return ""


def is_loose_output_name(value: str) -> bool:
    text = clean_text(value)
    if not text or TABLE_HEADER_PATTERN.search(text):
        return False
    if normalize_for_match(text) in {normalize_for_match(name) for name in GENERIC_STAGE_NAMES}:
        return False
    if looks_like_loose_output_id(text):
        return False
    return bool(re.search(r"[가-힣]", text))


def extract_runs(line: str) -> list[tuple[int, str]]:
    return [
        (match.start(), clean_text(match.group(0)))
        for match in RUN_PATTERN.finditer(line)
        if clean_text(match.group(0))
    ]


def candidate_output_name_after_id(after_parts: list[str]) -> str:
    # 관리문서 ID 표가 "ID 산출물명" 순서로 깨져 추출되는 경우 첫 번째 유효 컬럼을 산출물명으로 쓴다.
    for value in after_parts:
        candidate = clean_text(value)
        if not candidate or TABLE_HEADER_PATTERN.fullmatch(candidate):
            continue
        if OUTPUT_ID_IN_LINE_PATTERN.search(candidate):
            continue
        return candidate
    return ""


def aliases_for_output_name(
    output_name: str,
    id_name: str = "",
    table_output_name: str = "",
    trailing_output_name: str = "",
) -> tuple[str, ...]:
    # ID에 붙어 추출된 문서명과 산출물 컬럼 값을 모두 산출물명 후보로 보존한다.
    # 실제 파일 매칭은 ID가 아니라 이 산출물명 후보들을 기준으로 수행한다.
    extra_aliases = OUTPUT_NAME_ALIASES.get(output_name, ())
    if id_name:
        return unique_clean_values((output_name, id_name, *extra_aliases))
    return unique_clean_values((output_name, table_output_name, trailing_output_name, *extra_aliases))


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
