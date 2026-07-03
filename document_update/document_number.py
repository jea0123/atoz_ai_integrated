# 대상 한글/엑셀 파일 내부를 열어 문서번호, 프로젝트명, 문서제목을 실제로 치환한다.
# HWPX/Excel 파일 내부의 문서번호, 제목, 프로젝트명을 실제로 교체합니다.
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import shutil
import zipfile
from xml.sax.saxutils import escape, unescape

from .excel_ooxml import (
    EXCEL_DOCUMENT_SUFFIXES,
    build_updated_excel_cover_sheet,
    read_shared_strings,
    replace_excel_drawing_header_values,
    replace_excel_sheet_header_values,
    rewrite_excel_part_for_modified_workbook,
    workbook_sheets,
)
from .header_metadata import (
    unlabeled_header_metadata_indexes as shared_unlabeled_header_metadata_indexes,
)
from .hwpx_text import editable_hwpx_part_scope, is_hwpx_zip, split_hwpx_cover_edit_scope, strip_hwpx_line_seg_arrays
from .patterns import (
    CELL_PATTERN,
    OUTPUT_ID_PATTERN,
    ROW_PATTERN,
    RUN_OPEN_PATTERN,
    RUN_SELF_CLOSING_PATTERN,
    TEXT_NODE_PATTERN,
)
from .ppt_ooxml import PPT_DOCUMENT_SUFFIXES, write_updated_ppt_document
from .project_title_match import (
    best_matching_project_title,
    is_project_title_label_text,
)


DOCUMENT_NUMBER_LABEL = "\ubb38\uc11c\ubc88\ud638"
DOCUMENT_NUMBER_LABELS = {DOCUMENT_NUMBER_LABEL, "문서 번호", "문 서 번호"}
DOCUMENT_VERSION_VALUE = "v0.1"
UNLABELED_HEADER_VERSION_VALUE = DOCUMENT_VERSION_VALUE
DOCUMENT_VERSION_LABELS = {"문서버전", "문 서 버 전", "Version"}
DOCUMENT_NUMBER_LABEL_LIKE_VALUES = {
    "문서명",
    "문서제목",
    "산출물명",
    "문서버전",
    "버전",
    "개정일자",
    "작성자",
    "승인",
    "개정사유",
    "개정이력",
}
VERSION_PATTERN = re.compile(r"^[vV]?\d+(?:\.\d+)*$")
DATE_VALUE_PATTERN = re.compile(r"^\d{4}(?:[-./]\d{1,2}[-./]\d{1,2}|\s+\d{1,2}\s+\d{1,2})$")
HEADER_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?header)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
PARAGRAPH_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?p)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)


def normalize_document_number_label(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def is_document_number_label(value: str) -> bool:
    label = normalize_document_number_label(value.strip())
    return label in {normalize_document_number_label(item) for item in DOCUMENT_NUMBER_LABELS}


def is_document_version_label(value: str) -> bool:
    label = normalize_document_number_label(value.strip())
    return label in {normalize_document_number_label(item) for item in DOCUMENT_VERSION_LABELS}


def is_project_title_label(value: str) -> bool:
    return is_project_title_label_text(value)


def infer_document_number_from_filename(file_path: Path) -> str:
    """문서번호 셀이 비어 있을 때 대상 파일명에서 문서번호를 보조로 추정한다."""
    match = OUTPUT_ID_PATTERN.search(file_path.stem)
    if match:
        return match.group(0)

    return ""


def clean_text_node_value(value: str) -> str:
    return re.sub(r"\s+", " ", unescape(value)).strip()


def infer_cover_project_title_from_xml(xml: str, expected_project_title: str | None = None) -> str:
    """표지에서 사업명/프로젝트명 값을 보조 추정한다."""
    labeled_project_title = find_labeled_project_title_in_xml(xml)
    if labeled_project_title:
        return labeled_project_title

    text_values = [
        clean_text_node_value(match.group("body"))
        for match in TEXT_NODE_PATTERN.finditer(xml)
    ]
    for index, value in enumerate(text_values):
        if not is_document_number_label(value):
            continue

        candidates = [
            item
            for item in text_values[:index]
            if is_cover_title_value(item)
        ]
        if expected_project_title:
            return best_matching_project_title(candidates, expected_project_title)
        break

    return ""


def find_labeled_project_title_in_xml(xml: str) -> str:
    for row_match in ROW_PATTERN.finditer(xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))
        for cell_index, cell_match in enumerate(cells[:-1]):
            label = get_cell_text(cell_match.group(0)).strip()
            if not is_project_title_label(label):
                continue
            for target_cell in cells[cell_index + 1:]:
                value = get_cell_text(target_cell.group(0)).strip()
                if is_project_title_value(value):
                    return value

    text_values = [
        clean_text_node_value(match.group("body"))
        for match in TEXT_NODE_PATTERN.finditer(xml)
    ]
    for index, value in enumerate(text_values):
        match = re.search(r"^(?:사업명|프로젝트\s*명|프로젝트\s*제목)\s*[:：]\s*(.+)$", value)
        if match and is_project_title_value(match.group(1)):
            return clean_text_node_value(match.group(1))
        if is_project_title_label(value) and index + 1 < len(text_values):
            candidate = text_values[index + 1]
            if is_project_title_value(candidate):
                return candidate
    return ""


def is_project_title_value(value: str) -> bool:
    return is_cover_title_value(value) and not is_project_title_label(value)


def is_cover_title_value(value: str) -> bool:
    text = clean_text_node_value(value)
    if not text or len(text) > 120:
        return False
    if is_document_number_label(text) or is_document_version_label(text):
        return False
    if normalize_document_number_label(text) in {
        normalize_document_number_label(item)
        for item in DOCUMENT_NUMBER_LABEL_LIKE_VALUES
    }:
        return False
    if OUTPUT_ID_PATTERN.fullmatch(text):
        return False
    if VERSION_PATTERN.fullmatch(text) or DATE_VALUE_PATTERN.fullmatch(text):
        return False
    if text.startswith("㈜") or text.startswith("(주)"):
        return False
    return True


def unique_scan_values(*values: str) -> list[str]:
    result: list[str] = []
    for value in values:
        value = value.strip()
        if value and value not in result:
            result.append(value)

    return result


def append_unique_scan_values(values: list[str], *items: str) -> None:
    for item in items:
        item = item.strip()
        if item and item not in values:
            values.append(item)


def get_cell_text(cell_xml: str) -> str:
    return "".join(
        unescape(match.group("body"))
        for match in TEXT_NODE_PATTERN.finditer(cell_xml)
    )


def text_tag_for_run(run_tag: str) -> str:
    if ":" not in run_tag:
        return "t"

    return f"{run_tag.split(':', 1)[0]}:t"


def replace_cell_text(cell_xml: str, new_text: str) -> tuple[str, str]:
    """한글 확장 문서 표 셀 하나 안의 텍스트를 바꾸거나 새로 넣는다."""
    text_nodes = list(TEXT_NODE_PATTERN.finditer(cell_xml))
    old_text = get_cell_text(cell_xml)
    escaped_new_text = escape(new_text)

    if text_nodes:
        pieces: list[str] = []
        last = 0

        for index, match in enumerate(text_nodes):
            pieces.append(cell_xml[last:match.start("body")])
            pieces.append(escaped_new_text if index == 0 else "")
            last = match.end("body")

        pieces.append(cell_xml[last:])
        return "".join(pieces), old_text

    def replace_self_closing_run(match: re.Match[str]) -> str:
        text_tag = text_tag_for_run(match.group("tag"))
        return (
            f"<{match.group('tag')}{match.group('attrs')}>"
            f"<{text_tag}>{escaped_new_text}</{text_tag}>"
            f"</{match.group('tag')}>"
        )

    updated_cell, count = RUN_SELF_CLOSING_PATTERN.subn(
        replace_self_closing_run,
        cell_xml,
        count=1,
    )
    if count:
        return updated_cell, old_text

    run_match = RUN_OPEN_PATTERN.search(cell_xml)
    if run_match:
        text_tag = text_tag_for_run(run_match.group("tag"))
        insert_at = run_match.end()
        return (
            cell_xml[:insert_at]
            + f"<{text_tag}>{escaped_new_text}</{text_tag}>"
            + cell_xml[insert_at:]
        ), old_text

    paragraph_close = re.search(r"</(?:\w+:)?p>", cell_xml)
    if paragraph_close:
        insert_at = paragraph_close.start()
        return (
            cell_xml[:insert_at]
            + f"<hp:run><hp:t>{escaped_new_text}</hp:t></hp:run>"
            + cell_xml[insert_at:]
        ), old_text

    raise RuntimeError("문서번호 오른쪽 칸에 텍스트를 넣을 위치를 찾지 못했습니다.")


def find_document_number_cell_value(xml: str) -> str:
    for row_match in ROW_PATTERN.finditer(xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))
        for cell_index, cell_match in enumerate(cells[:-1]):
            label_text = get_cell_text(cell_match.group(0)).strip()
            if is_document_number_label(label_text):
                target_text = get_cell_text(cells[cell_index + 1].group(0)).strip()
                if target_text and not is_document_number_value_cell(target_text):
                    continue
                return target_text
    return ""


def is_document_number_value_cell(value: str) -> bool:
    text = value.strip()
    if not text:
        return False
    label_like_values = {normalize_document_number_label(item) for item in DOCUMENT_NUMBER_LABEL_LIKE_VALUES}
    return normalize_document_number_label(text) not in label_like_values


def is_document_version_value_cell(value: str) -> bool:
    label_like_values = {normalize_document_number_label(item) for item in DOCUMENT_NUMBER_LABEL_LIKE_VALUES}
    return normalize_document_number_label(value) not in label_like_values


def replace_document_number_labeled_cells(
    xml: str,
    new_document_number: str,
) -> tuple[str, list[str], int, int]:
    pieces: list[str] = []
    last = 0
    old_values: list[str] = []
    changed_count = 0
    found_count = 0

    for row_match in ROW_PATTERN.finditer(xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))
        updated_row = row_xml
        offset = 0
        row_changed = False

        for cell_index, cell_match in enumerate(cells[:-1]):
            label_text = get_cell_text(cell_match.group(0)).strip()
            if not is_document_number_label(label_text):
                continue

            target_cell_match = cells[cell_index + 1]
            old_document_number = get_cell_text(target_cell_match.group(0)).strip()
            if old_document_number and not is_document_number_value_cell(old_document_number):
                continue
            old_values.append(old_document_number)
            found_count += 1
            if old_document_number == new_document_number:
                continue

            updated_cell, _old_document_number = replace_cell_text(
                target_cell_match.group(0),
                new_document_number,
            )
            start = target_cell_match.start() + offset
            end = target_cell_match.end() + offset
            updated_row = updated_row[:start] + updated_cell + updated_row[end:]
            offset += len(updated_cell) - len(target_cell_match.group(0))
            changed_count += 1
            row_changed = True

        if row_changed:
            pieces.append(xml[last:row_match.start()])
            pieces.append(updated_row)
            last = row_match.end()

    if not pieces:
        return xml, old_values, changed_count, found_count

    pieces.append(xml[last:])
    return "".join(pieces), old_values, changed_count, found_count


def replace_document_version_labeled_cells(xml: str) -> tuple[str, int]:
    pieces: list[str] = []
    last = 0
    changed_count = 0

    for row_match in ROW_PATTERN.finditer(xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))
        updated_row = row_xml
        offset = 0
        row_changed = False

        for cell_index, cell_match in enumerate(cells[:-1]):
            label_text = get_cell_text(cell_match.group(0)).strip()
            if not is_document_version_label(label_text):
                continue

            target_cell_match = cells[cell_index + 1]
            old_version = get_cell_text(target_cell_match.group(0)).strip()
            if not is_document_version_value_cell(old_version) or old_version == DOCUMENT_VERSION_VALUE:
                continue

            updated_cell, _old_version = replace_cell_text(target_cell_match.group(0), DOCUMENT_VERSION_VALUE)
            start = target_cell_match.start() + offset
            end = target_cell_match.end() + offset
            updated_row = updated_row[:start] + updated_cell + updated_row[end:]
            offset += len(updated_cell) - len(target_cell_match.group(0))
            changed_count += 1
            row_changed = True

        if row_changed:
            pieces.append(xml[last:row_match.start()])
            pieces.append(updated_row)
            last = row_match.end()

    if not pieces:
        return xml, 0

    pieces.append(xml[last:])
    return "".join(pieces), changed_count


def replace_unlabeled_header_version_cells(xml: str) -> tuple[str, int]:
    pieces: list[str] = []
    last = 0
    changed_count = 0

    for header_match in HEADER_PATTERN.finditer(xml):
        header_xml = header_match.group(0)
        updated_header, count = replace_unlabeled_header_version_block(header_xml)
        if not count:
            continue
        pieces.append(xml[last:header_match.start()])
        pieces.append(updated_header)
        last = header_match.end()
        changed_count += count

    if not pieces:
        return xml, 0

    pieces.append(xml[last:])
    return "".join(pieces), changed_count


def replace_unlabeled_header_version_block(header_xml: str) -> tuple[str, int]:
    for row_match in ROW_PATTERN.finditer(header_xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))
        if len(cells) < 4:
            continue

        values = [get_cell_text(cell.group(0)).strip() for cell in cells]
        indexes = unlabeled_header_metadata_indexes(values)
        if indexes is None:
            continue

        _code_index, version_index, _date_index, _author_index = indexes
        version_cell = cells[version_index]
        old_version = values[version_index]
        if old_version == UNLABELED_HEADER_VERSION_VALUE:
            return header_xml, 0

        updated_cell, _old_version = replace_cell_text(version_cell.group(0), UNLABELED_HEADER_VERSION_VALUE)
        updated_row = row_xml[:version_cell.start()] + updated_cell + row_xml[version_cell.end():]
        return header_xml[:row_match.start()] + updated_row + header_xml[row_match.end():], 1

    return header_xml, 0


def unlabeled_header_metadata_indexes(values: list[str]) -> tuple[int, int, int, int] | None:
    return shared_unlabeled_header_metadata_indexes(
        values,
        clean_text=clean_text_node_value,
        normalize_label=normalize_document_number_label,
        label_like_values=DOCUMENT_NUMBER_LABEL_LIKE_VALUES,
        version_pattern=VERSION_PATTERN,
        date_pattern=DATE_VALUE_PATTERN,
    )


def replace_matching_text_nodes(xml: str, old_text: str | None, new_text: str | None) -> tuple[str, int]:
    """머리글 같은 곳에 반복된 동일 텍스트 노드를 정확히 일치할 때만 바꾼다."""
    if not old_text or not new_text or old_text == new_text:
        return xml, 0

    escaped_new_text = escape(new_text)
    pieces: list[str] = []
    last = 0
    replace_count = 0

    for match in TEXT_NODE_PATTERN.finditer(xml):
        body = match.group("body")
        if unescape(body).strip() != old_text:
            continue

        pieces.append(xml[last:match.start("body")])
        pieces.append(escaped_new_text)
        last = match.end("body")
        replace_count += 1

    if replace_count == 0:
        return replace_matching_paragraph_text(xml, old_text, new_text)

    pieces.append(xml[last:])
    return "".join(pieces), replace_count


def replace_matching_paragraph_text(xml: str, old_text: str, new_text: str) -> tuple[str, int]:
    escaped_new_text = escape(new_text)
    pieces: list[str] = []
    last = 0
    replace_count = 0

    for paragraph_match in PARAGRAPH_PATTERN.finditer(xml):
        paragraph_xml = paragraph_match.group(0)
        text_nodes = list(TEXT_NODE_PATTERN.finditer(paragraph_xml))
        if not text_nodes:
            continue

        combined = "".join(unescape(match.group("body")) for match in text_nodes).strip()
        if combined != old_text:
            continue

        paragraph_pieces: list[str] = []
        paragraph_last = 0
        for index, text_match in enumerate(text_nodes):
            paragraph_pieces.append(paragraph_xml[paragraph_last:text_match.start("body")])
            paragraph_pieces.append(escaped_new_text if index == 0 else "")
            paragraph_last = text_match.end("body")
        paragraph_pieces.append(paragraph_xml[paragraph_last:])

        pieces.append(xml[last:paragraph_match.start()])
        pieces.append("".join(paragraph_pieces))
        last = paragraph_match.end()
        replace_count += 1

    if replace_count == 0:
        return xml, 0

    pieces.append(xml[last:])
    return "".join(pieces), replace_count


def backup_path_for(file_path: Path) -> Path:
    return file_path.parent / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_path.suffix}"


def write_updated_document(
    file_path: Path,
    new_document_number: str,
    old_project_title: str | None = None,
    new_project_title: str | None = None,
    output_path: Path | None = None,
    *,
    allow_missing_document_number: bool = False,
) -> tuple[str, Path, int, int, Path]:
    """대상 파일 형식에 맞는 수정 함수로 분기한다."""
    suffix = file_path.suffix.lower()

    if is_hwpx_zip(file_path):
        return write_updated_hwpx_document(
            file_path,
            new_document_number,
            old_project_title,
            new_project_title,
            output_path,
            allow_missing_document_number=allow_missing_document_number,
        )

    if suffix in EXCEL_DOCUMENT_SUFFIXES:
        return write_updated_excel_document(
            file_path,
            new_document_number,
            old_project_title,
            new_project_title,
            output_path,
            allow_missing_document_number=allow_missing_document_number,
        )

    if suffix in PPT_DOCUMENT_SUFFIXES:
        return write_updated_ppt_document(
            file_path,
            new_document_number,
            old_project_title,
            new_project_title,
            output_path,
            allow_missing_document_number=allow_missing_document_number,
        )

    raise RuntimeError(
        "지원하지 않는 대상 파일 형식입니다. "
        "대상 파일은 HWP/HWPX, XLSX/XLSM 또는 PPT/PPTX 계열을 사용해주세요."
    )


def write_updated_project_title(
    file_path: Path,
    old_project_title: str,
    new_project_title: str,
    output_path: Path | None = None,
) -> tuple[int, Path]:
    """표준 산출물 매칭 없이 표지/문서 안의 기존 프로젝트명만 교체한다."""
    old_project_title = clean_text_node_value(old_project_title)
    new_project_title = clean_text_node_value(new_project_title)
    if not old_project_title or not new_project_title or old_project_title == new_project_title:
        return 0, file_path

    suffix = file_path.suffix.lower()
    if not (is_hwpx_zip(file_path) or suffix in EXCEL_DOCUMENT_SUFFIXES or suffix in PPT_DOCUMENT_SUFFIXES):
        raise RuntimeError("프로젝트명만 교체할 수 있는 지원 형식이 아닙니다.")

    backup_path = backup_path_for(file_path)
    write_path = output_path or file_path.parent / f"working_project_title{file_path.suffix}"
    if output_path is None:
        shutil.copy2(file_path, backup_path)

    replace_count = 0
    is_hwpx = is_hwpx_zip(file_path)
    try:
        with zipfile.ZipFile(file_path, "r") as zin:
            with zipfile.ZipFile(write_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.lower().endswith(".xml"):
                        xml = data.decode("utf-8", errors="ignore")
                        editable_xml, preserved_xml = (
                            editable_hwpx_part_scope(item.filename, xml) if is_hwpx else (xml, "")
                        )
                        if not editable_xml:
                            zout.writestr(item, data)
                            continue
                        editable_xml, count = replace_matching_text_nodes(editable_xml, old_project_title, new_project_title)
                        if count:
                            replace_count += count
                            if is_hwpx:
                                editable_xml, _line_seg_count = strip_hwpx_line_seg_arrays(editable_xml)
                            data = (editable_xml + preserved_xml).encode("utf-8")
                    zout.writestr(item, data)

        updated_path = write_path
        if output_path is None:
            write_path.replace(file_path)
            updated_path = file_path
        return replace_count, updated_path
    except Exception:
        if output_path is None and write_path.exists():
            try:
                write_path.unlink()
            except OSError:
                pass
        raise


# 한글 확장 문서는 압축 파일 안에 문서 조각이 들어 있으므로 내부 텍스트 노드를 수정한다.
def write_updated_hwpx_document(
    file_path: Path,
    new_document_number: str,
    old_project_title: str | None = None,
    new_project_title: str | None = None,
    output_path: Path | None = None,
    *,
    allow_missing_document_number: bool = False,
) -> tuple[str, Path, int, int, Path]:
    """업로드된 대상 파일을 기준으로 수정된 한글 확장 결과 파일을 만든다."""
    backup_path = backup_path_for(file_path)
    write_path = output_path or file_path.parent / "working_output.hwpx"

    with zipfile.ZipFile(file_path, "r") as source_zip:
        section_xml = source_zip.read("Contents/section0.xml").decode("utf-8", errors="ignore")
        cover_xml, _body_xml = split_hwpx_cover_edit_scope(section_xml)
        old_document_number = find_document_number_cell_value(cover_xml)
        inferred_project_title = infer_cover_project_title_from_xml(
            cover_xml,
            expected_project_title=new_project_title,
        )

    document_numbers_to_scan = unique_scan_values(
        old_document_number,
        infer_document_number_from_filename(file_path),
    )
    project_titles_to_scan = unique_scan_values(
        old_project_title or "",
        inferred_project_title,
    )

    if output_path is None:
        # 원본을 덮어쓸 때는 먼저 시각이 들어간 백업을 만든다.
        shutil.copy2(file_path, backup_path)

    project_title_replace_count = 0
    matching_document_number_replace_count = 0
    document_number_position_found = False

    try:
        with zipfile.ZipFile(file_path, "r") as zin:
            with zipfile.ZipFile(write_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename.lower().endswith(".xml"):
                        xml = data.decode("utf-8", errors="ignore")
                        editable_xml, preserved_xml = editable_hwpx_part_scope(item.filename, xml)
                        if not editable_xml:
                            zout.writestr(item, data)
                            continue
                        changed = False

                        for project_title_to_scan in project_titles_to_scan:
                            editable_xml, count = replace_matching_text_nodes(
                                editable_xml,
                                project_title_to_scan,
                                new_project_title,
                            )
                            if count:
                                project_title_replace_count += count
                                changed = True

                        (
                            editable_xml,
                            labeled_old_document_numbers,
                            labeled_replace_count,
                            labeled_found_count,
                        ) = replace_document_number_labeled_cells(editable_xml, new_document_number)
                        if labeled_found_count:
                            document_number_position_found = True
                        if labeled_replace_count:
                            append_unique_scan_values(document_numbers_to_scan, *labeled_old_document_numbers)
                            matching_document_number_replace_count += labeled_replace_count
                            changed = True

                        editable_xml, count = replace_document_version_labeled_cells(editable_xml)
                        if count:
                            changed = True

                        editable_xml, count = replace_unlabeled_header_version_cells(editable_xml)
                        if count:
                            changed = True

                        for document_number_to_scan in document_numbers_to_scan:
                            editable_xml, count = replace_matching_text_nodes(
                                editable_xml,
                                document_number_to_scan,
                                new_document_number,
                            )
                            if count:
                                matching_document_number_replace_count += count
                                changed = True

                        if changed:
                            editable_xml, _line_seg_count = strip_hwpx_line_seg_arrays(editable_xml)
                            data = (editable_xml + preserved_xml).encode("utf-8")

                    zout.writestr(item, data)

        if (
            not allow_missing_document_number
            and not document_number_position_found
            and not matching_document_number_replace_count
        ):
            raise RuntimeError("첫 번째 장에서 문서번호 바로 오른쪽 ID 칸을 찾지 못했습니다.")

        updated_path = write_path
        if output_path is None:
            write_path.replace(file_path)
            updated_path = file_path

        return (
            old_document_number,
            backup_path,
            project_title_replace_count,
            matching_document_number_replace_count,
            updated_path,
        )

    except Exception:
        if output_path is None and write_path.exists():
            try:
                write_path.unlink()
            except OSError:
                pass
        raise


def write_updated_excel_document(
    file_path: Path,
    new_document_number: str,
    old_project_title: str | None = None,
    new_project_title: str | None = None,
    output_path: Path | None = None,
    *,
    allow_missing_document_number: bool = False,
) -> tuple[str, Path, int, int, Path]:
    """엑셀 통합문서의 표지 시트 셀을 수정한다."""
    backup_path = backup_path_for(file_path)
    write_path = output_path or file_path.parent / f"working_output{file_path.suffix}"

    if output_path is None:
        shutil.copy2(file_path, backup_path)

    try:
        with zipfile.ZipFile(file_path, "r") as zin:
            shared_strings = read_shared_strings(zin)
            (
                cover_sheet_path,
                updated_cover_sheet,
                old_document_number,
                project_title_replace_count,
                document_number_replace_count,
                project_titles_to_scan,
            ) = build_updated_excel_cover_sheet(
                zin,
                new_document_number,
                old_project_title,
                new_project_title,
                allow_missing_document_number=allow_missing_document_number,
            )
            document_numbers_to_scan = unique_scan_values(
                old_document_number,
                infer_document_number_from_filename(file_path),
            )
            updated_sheets: dict[str, bytes] = {}

            for sheet in workbook_sheets(zin):
                sheet_xml = (
                    updated_cover_sheet.decode("utf-8", errors="ignore")
                    if sheet.path == cover_sheet_path
                    else zin.read(sheet.path).decode("utf-8", errors="ignore")
                )
                (
                    updated_sheet,
                    header_project_count,
                    header_document_number_count,
                    header_version_count,
                ) = replace_excel_sheet_header_values(
                    sheet_xml,
                    shared_strings,
                    project_titles_to_scan,
                    new_project_title,
                    document_numbers_to_scan,
                    new_document_number,
                )
                if sheet.path == cover_sheet_path or header_project_count or header_document_number_count or header_version_count:
                    updated_sheets[sheet.path] = updated_sheet
                    project_title_replace_count += header_project_count
                    document_number_replace_count += header_document_number_count

            with zipfile.ZipFile(write_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename in updated_sheets:
                        data = updated_sheets[item.filename]
                    elif item.filename.startswith("xl/drawings/") and item.filename.endswith(".xml"):
                        (
                            data,
                            drawing_project_count,
                            drawing_document_number_count,
                        ) = replace_excel_drawing_header_values(
                            data.decode("utf-8", errors="ignore"),
                            project_titles_to_scan,
                            new_project_title,
                            document_numbers_to_scan,
                            new_document_number,
                        )
                        project_title_replace_count += drawing_project_count
                        document_number_replace_count += drawing_document_number_count
                    rewritten = rewrite_excel_part_for_modified_workbook(item.filename, data)
                    if rewritten is None:
                        continue
                    data = rewritten
                    zout.writestr(item, data)

        updated_path = write_path
        if output_path is None:
            write_path.replace(file_path)
            updated_path = file_path

        return (
            old_document_number,
            backup_path,
            project_title_replace_count,
            document_number_replace_count,
            updated_path,
        )
    except Exception:
        if output_path is None and write_path.exists():
            try:
                write_path.unlink()
            except OSError:
                pass
        raise
