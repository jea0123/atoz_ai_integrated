# 엑셀 문서 구조의 표지 시트에서 제목/프로젝트명/문서번호 셀을 읽고 수정한다.
# Excel OOXML 구조를 직접 읽고 표지 시트의 제목/프로젝트명/문서번호를 수정합니다.
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path, PurePosixPath
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, unescape

from .patterns import OUTPUT_ID_PATTERN


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"
DOCUMENT_NUMBER_LABEL = "\ubb38\uc11c\ubc88\ud638"
DOCUMENT_VERSION_VALUE = "v0.1"
UNLABELED_HEADER_VERSION_VALUE = DOCUMENT_VERSION_VALUE
DOCUMENT_VERSION_LABELS = {"문서버전", "문 서 버 전", "Version"}
LABEL_LIKE_VALUES = {
    "문서명",
    "문서제목",
    "산출물명",
    "문서번호",
    "문서버전",
    "버전",
    "개정일자",
    "작성자",
    "승인",
    "개정사유",
    "개정이력",
}
EXCEL_DOCUMENT_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
COVER_SHEET_HINT = "\ud45c\uc9c0"
DRAWING_PARAGRAPH_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?p)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
DRAWING_TEXT_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?t)\b[^>]*>(?P<text>.*?)</(?P=tag)>", re.DOTALL)
VERSION_PATTERN = re.compile(r"^[vV]?\d+(?:\.\d+)*$")
DATE_VALUE_PATTERN = re.compile(r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}$")

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def qn(local_name: str) -> str:
    return f"{{{MAIN_NS}}}{local_name}"


@dataclass(frozen=True)
class WorksheetRef:
    name: str
    path: str


@dataclass(frozen=True)
class CellInfo:
    ref: str
    row: int
    col: int
    text: str
    cell: ET.Element
    row_element: ET.Element


def normalize_part_path(target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return str(PurePosixPath(target.lstrip("/")))
    if target.startswith("xl/"):
        return str(PurePosixPath(target))
    return str(PurePosixPath("xl") / target)


def workbook_sheets(zf: zipfile.ZipFile) -> list[WorksheetRef]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }

    sheets_element = workbook.find(qn("sheets"))
    if sheets_element is None:
        return []

    result: list[WorksheetRef] = []
    for sheet in sheets_element:
        rel_id = sheet.attrib.get(f"{{{REL_NS}}}id", "")
        target = rel_map.get(rel_id, "")
        if not target:
            continue
        result.append(
            WorksheetRef(
                name=sheet.attrib.get("name", ""),
                path=normalize_part_path(target),
            )
        )
    return result


def cover_sheet_ref(zf: zipfile.ZipFile) -> WorksheetRef:
    sheets = workbook_sheets(zf)
    if not sheets:
        raise RuntimeError("엑셀 통합문서에서 시트를 찾지 못했습니다.")

    for sheet in sheets:
        if COVER_SHEET_HINT in sheet.name:
            return sheet
    return sheets[0]


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for item in root.findall(qn("si")):
        values.append("".join(text.text or "" for text in item.iter(qn("t"))))
    return values


def cell_ref_to_col(ref: str) -> int:
    letters = "".join(ch for ch in ref if ch.isalpha())
    value = 0
    for letter in letters:
        value = value * 26 + (ord(letter.upper()) - 64)
    return value


def col_to_name(col: int) -> str:
    letters: list[str] = []
    while col:
        col, remainder = divmod(col - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        value = cell.find(qn("v"))
        if value is None or value.text is None:
            return ""
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""

    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(qn("t")))

    value = cell.find(qn("v"))
    return value.text if value is not None and value.text is not None else ""


def row_number(row: ET.Element) -> int:
    raw = row.attrib.get("r")
    if raw and raw.isdigit():
        return int(raw)

    for cell in row.findall(qn("c")):
        match = re.search(r"\d+", cell.attrib.get("r", ""))
        if match:
            return int(match.group(0))

    return 0


def iter_cells(root: ET.Element, shared_strings: list[str]) -> list[CellInfo]:
    sheet_data = root.find(qn("sheetData"))
    if sheet_data is None:
        return []

    cells: list[CellInfo] = []
    for row in sheet_data.findall(qn("row")):
        current_row = row_number(row)
        for cell in row.findall(qn("c")):
            ref = cell.attrib.get("r", "")
            if not ref:
                continue
            cells.append(
                CellInfo(
                    ref=ref,
                    row=current_row,
                    col=cell_ref_to_col(ref),
                    text=cell_text(cell, shared_strings),
                    cell=cell,
                    row_element=row,
                )
            )
    return cells


def clean_cover_value(value: str) -> str:
    return value.strip().strip("\"'`<>")


def normalize_metadata_label(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def is_document_version_label(value: str) -> bool:
    label = normalize_metadata_label(clean_cover_value(value))
    return label in {normalize_metadata_label(item) for item in DOCUMENT_VERSION_LABELS}


def is_label_like_value(value: str) -> bool:
    label = normalize_metadata_label(clean_cover_value(value))
    return label in {normalize_metadata_label(item) for item in LABEL_LIKE_VALUES}


def meaningful_cover_value(value: str) -> str:
    cleaned = clean_cover_value(value)
    if not cleaned:
        return ""

    noise_prefixes = (
        "\uc2dd\ud488\uc758\uc57d\ud488\uc548\uc804\ucc98",
        "\u321c",
        "(\uc8fc)",
        "\uac1c \uc815 \uc774 \ub825",
        "\ubaa9 \ucc28",
    )
    noise_values = {
        DOCUMENT_NUMBER_LABEL,
        "\ubb38\uc11c\ubc84\uc804",
        "\uac1c\uc815\uc77c\uc790",
        "\uc791  \uc131 \uc790",
        "\uc791\uc131\uc790",
    }

    if cleaned in noise_values:
        return ""
    if any(cleaned.startswith(prefix) for prefix in noise_prefixes):
        return ""
    if OUTPUT_ID_PATTERN.fullmatch(cleaned):
        return ""

    return cleaned


def find_excel_cover_identity(file_path: Path) -> tuple[str, str]:
    with zipfile.ZipFile(file_path, "r") as zf:
        shared_strings = read_shared_strings(zf)
        sheet = cover_sheet_ref(zf)
        root = ET.fromstring(zf.read(sheet.path))
        cells = iter_cells(root, shared_strings)

    document_rows = [
        cell.row
        for cell in cells
        if clean_cover_value(cell.text) == DOCUMENT_NUMBER_LABEL
    ]
    cutoff_row = min(document_rows) if document_rows else 40

    values: list[str] = []
    for cell in sorted(cells, key=lambda item: (item.row, item.col)):
        if cell.row >= cutoff_row:
            continue
        value = meaningful_cover_value(cell.text)
        if value and value not in values:
            values.append(value)

    if len(values) >= 2:
        return values[-2], values[-1]
    return "", ""


def extract_excel_cover_text(file_path: Path, max_chars: int = 1000) -> str:
    """표지/첫 워크시트의 앞쪽 셀만 행/열 순서로 읽는다."""
    with zipfile.ZipFile(file_path, "r") as zf:
        shared_strings = read_shared_strings(zf)
        sheet = cover_sheet_ref(zf)
        root = ET.fromstring(zf.read(sheet.path))
        cells = iter_cells(root, shared_strings)

    document_rows = [
        cell.row
        for cell in cells
        if clean_cover_value(cell.text) == DOCUMENT_NUMBER_LABEL
    ]
    cutoff_row = min(document_rows) + 8 if document_rows else 35

    values: list[str] = []
    for cell in sorted(cells, key=lambda item: (item.row, item.col)):
        if cell.row > cutoff_row:
            continue
        value = clean_cover_value(cell.text)
        if value and value not in values:
            values.append(value)
        if len(" ".join(values)) >= max_chars:
            break

    return " ".join(values)[:max_chars]


def element_prefix(tag: str) -> str:
    return f"{tag.split(':', 1)[0]}:" if ":" in tag else ""


def cell_xml_pattern(cell_ref: str) -> re.Pattern[str]:
    return re.compile(
        rf"<(?P<tag>(?:\w+:)?c)\b(?=[^>]*\br=\"{re.escape(cell_ref)}\")[^>]*(?:/>|>.*?</(?P=tag)>)",
        re.DOTALL,
    )


def row_xml_pattern(row_index: int) -> re.Pattern[str]:
    return re.compile(
        rf"<(?P<tag>(?:\w+:)?row)\b(?=[^>]*\br=\"{row_index}\")[^>]*(?:/>|>.*?</(?P=tag)>)",
        re.DOTALL,
    )


def any_cell_xml_pattern() -> re.Pattern[str]:
    return re.compile(r"<(?P<tag>(?:\w+:)?c)\b[^>]*(?:/>|>.*?</(?P=tag)>)", re.DOTALL)


def ref_from_cell_xml(cell_xml: str) -> str:
    match = re.search(r'\br="([^"]+)"', cell_xml)
    return match.group(1) if match else ""


def inline_text_payload(prefix: str, new_text: str) -> str:
    space_attr = ' xml:space="preserve"' if new_text != new_text.strip() else ""
    escaped_text = escape(new_text)
    return f"<{prefix}is><{prefix}t{space_attr}>{escaped_text}</{prefix}t></{prefix}is>"


def inline_cell_xml(cell_ref: str, new_text: str, prefix: str = "") -> str:
    return f'<{prefix}c r="{cell_ref}" t="inlineStr">{inline_text_payload(prefix, new_text)}</{prefix}c>'


def replace_cell_xml(cell_xml: str, new_text: str) -> str:
    start_match = re.match(r"<(?P<tag>(?:\w+:)?c)\b(?P<attrs>[^>]*)/?>?", cell_xml, re.DOTALL)
    if not start_match:
        raise RuntimeError("엑셀 셀 XML을 해석하지 못했습니다.")

    tag = start_match.group("tag")
    attrs = re.sub(r'\s+t="[^"]*"', "", start_match.group("attrs"))
    prefix = element_prefix(tag)
    return f"<{tag}{attrs} t=\"inlineStr\">{inline_text_payload(prefix, new_text)}</{tag}>"


def replace_or_insert_cell_xml(
    sheet_xml: str,
    cell_ref: str,
    row_index: int,
    col_index: int,
    new_text: str,
) -> str:
    """시트를 다시 직렬화하지 않고 원본 워크시트 문서 조각에서 셀 하나만 수정한다."""
    cell_match = cell_xml_pattern(cell_ref).search(sheet_xml)
    if cell_match:
        return (
            sheet_xml[:cell_match.start()]
            + replace_cell_xml(cell_match.group(0), new_text)
            + sheet_xml[cell_match.end():]
        )

    row_match = row_xml_pattern(row_index).search(sheet_xml)
    if not row_match:
        raise RuntimeError(f"엑셀 표지에서 {row_index}행을 찾지 못했습니다.")

    row_xml = row_match.group(0)
    row_tag = row_match.group("tag")
    prefix = element_prefix(row_tag)
    new_cell = inline_cell_xml(cell_ref, new_text, prefix)

    if row_xml.endswith("/>"):
        open_row = row_xml[:-2] + ">"
        updated_row = f"{open_row}{new_cell}</{row_tag}>"
    else:
        inserted = False
        updated_row_parts: list[str] = []
        last = 0
        for candidate in any_cell_xml_pattern().finditer(row_xml):
            candidate_ref = ref_from_cell_xml(candidate.group(0))
            if candidate_ref and cell_ref_to_col(candidate_ref) > col_index:
                updated_row_parts.append(row_xml[last:candidate.start()])
                updated_row_parts.append(new_cell)
                updated_row_parts.append(row_xml[candidate.start():])
                inserted = True
                break

        if inserted:
            updated_row = "".join(updated_row_parts)
        else:
            close_match = re.search(rf"</{re.escape(row_tag)}>\s*$", row_xml)
            if not close_match:
                raise RuntimeError(f"엑셀 표지에서 {row_index}행 닫는 태그를 찾지 못했습니다.")
            updated_row = row_xml[:close_match.start()] + new_cell + row_xml[close_match.start():]

    return sheet_xml[:row_match.start()] + updated_row + sheet_xml[row_match.end():]


def build_updated_excel_cover_sheet(
    zf: zipfile.ZipFile,
    new_document_number: str,
    old_title: str | None,
    new_title: str | None,
    old_project_title: str | None,
    new_project_title: str | None,
) -> tuple[str, bytes, str, int, int, int]:
    shared_strings = read_shared_strings(zf)
    sheet = cover_sheet_ref(zf)
    original_sheet_xml = zf.read(sheet.path).decode("utf-8", errors="ignore")
    root = ET.fromstring(original_sheet_xml)

    title_count = 0
    project_count = 0
    old_document_number = ""
    document_number_count = 0
    updates: dict[str, tuple[int, int, str]] = {}

    for info in iter_cells(root, shared_strings):
        value = clean_cover_value(info.text)
        if old_title and new_title and old_title != new_title and value == old_title:
            updates[info.ref] = (info.row, info.col, new_title)
            title_count += 1
        if (
            old_project_title
            and new_project_title
            and old_project_title != new_project_title
            and value == old_project_title
        ):
            updates[info.ref] = (info.row, info.col, new_project_title)
            project_count += 1

    cell_by_ref = {info.ref: info for info in iter_cells(root, shared_strings)}
    for info in iter_cells(root, shared_strings):
        if clean_cover_value(info.text) != DOCUMENT_NUMBER_LABEL:
            continue

        target_ref = f"{col_to_name(info.col + 1)}{info.row}"
        target_info = cell_by_ref.get(target_ref)
        old_document_number = clean_cover_value(target_info.text if target_info else "")
        if old_document_number != new_document_number:
            updates[target_ref] = (info.row, info.col + 1, new_document_number)
        document_number_count = 0 if old_document_number == new_document_number else 1
        break

    cell_by_position = {(info.row, info.col): info for info in iter_cells(root, shared_strings)}
    for info in iter_cells(root, shared_strings):
        if not is_document_version_label(info.text):
            continue

        target_col = info.col + 1
        target_ref = f"{col_to_name(target_col)}{info.row}"
        target_info = cell_by_position.get((info.row, target_col))
        old_version = clean_cover_value(target_info.text if target_info else "")
        if is_label_like_value(old_version) or old_version == DOCUMENT_VERSION_VALUE:
            continue
        updates[target_ref] = (info.row, target_col, DOCUMENT_VERSION_VALUE)

    if not old_document_number and document_number_count == 0:
        raise RuntimeError("엑셀 표지에서 문서번호 오른쪽 칸을 찾지 못했습니다.")

    updated_xml_text = original_sheet_xml
    for cell_ref, (row_index, col_index, value) in updates.items():
        updated_xml_text = replace_or_insert_cell_xml(
            updated_xml_text,
            cell_ref,
            row_index,
            col_index,
            value,
        )

    return (
        sheet.path,
        updated_xml_text.encode("utf-8"),
        old_document_number,
        title_count,
        project_count,
        document_number_count,
    )


EXCEL_HEADER_SCAN_MAX_ROW = 8


def replace_excel_sheet_header_values(
    sheet_xml: str,
    shared_strings: list[str],
    old_project_title: str | None,
    new_project_title: str | None,
    old_document_numbers: list[str],
    new_document_number: str,
) -> tuple[bytes, int, int, int]:
    root = ET.fromstring(sheet_xml)
    replacements: dict[str, tuple[str, str]] = {}
    cells = iter_cells(root, shared_strings)

    old_project = clean_cover_value(old_project_title or "")
    new_project = clean_cover_value(new_project_title or "")
    if old_project and new_project and old_project != new_project:
        replacements[old_project] = (new_project_title or "", "project")

    for old_document_number in old_document_numbers:
        old_document = clean_cover_value(old_document_number)
        if old_document and old_document != clean_cover_value(new_document_number):
            replacements[old_document] = (new_document_number, "document")

    updates: dict[str, tuple[int, int, str, str]] = {}
    cells_by_position = {(cell.row, cell.col): cell for cell in cells}
    for info in cells:
        if info.row > EXCEL_HEADER_SCAN_MAX_ROW:
            continue
        replacement = replacements.get(clean_cover_value(info.text))
        if replacement:
            updates[info.ref] = (info.row, info.col, replacement[0], replacement[1])
        if is_document_version_label(info.text):
            target_col = info.col + 1
            target_ref = f"{col_to_name(target_col)}{info.row}"
            target_info = cells_by_position.get((info.row, target_col))
            old_version = clean_cover_value(target_info.text if target_info else "")
            if not is_label_like_value(old_version) and old_version != DOCUMENT_VERSION_VALUE:
                updates[target_ref] = (info.row, target_col, DOCUMENT_VERSION_VALUE, "version")

    for cell_ref, row, col, current_value in unlabeled_header_version_targets(cells):
        if clean_cover_value(current_value) != UNLABELED_HEADER_VERSION_VALUE:
            updates[cell_ref] = (
                row,
                col,
                UNLABELED_HEADER_VERSION_VALUE,
                "version",
            )

    if not updates:
        return sheet_xml.encode("utf-8"), 0, 0, 0

    updated_xml_text = sheet_xml
    project_count = 0
    document_number_count = 0
    version_count = 0
    for cell_ref, (row_index, col_index, value, kind) in updates.items():
        updated_xml_text = replace_or_insert_cell_xml(
            updated_xml_text,
            cell_ref,
            row_index,
            col_index,
            value,
        )
        if kind == "project":
            project_count += 1
        elif kind == "document":
            document_number_count += 1
        elif kind == "version":
            version_count += 1

    return updated_xml_text.encode("utf-8"), project_count, document_number_count, version_count


def unlabeled_header_version_targets(cells: list[CellInfo]) -> list[tuple[str, int, int, str]]:
    rows: dict[int, list[CellInfo]] = {}
    for cell in cells:
        if cell.row <= EXCEL_HEADER_SCAN_MAX_ROW:
            rows.setdefault(cell.row, []).append(cell)

    result: list[tuple[str, int, int, str]] = []
    for row_index in sorted(rows):
        row_cells = sorted(rows[row_index], key=lambda item: item.col)
        row_values = [clean_cover_value(cell.text) for cell in row_cells]
        output_cell = next((cell for cell in row_cells if OUTPUT_ID_PATTERN.search(clean_cover_value(cell.text))), None)
        if output_cell is None:
            continue

        date_cell = next(
            (cell for cell in row_cells if DATE_VALUE_PATTERN.fullmatch(clean_cover_value(cell.text))),
            None,
        )
        if date_cell is None:
            continue

        version_cell = next(
            (
                cell
                for cell in row_cells
                if cell.col < date_cell.col and VERSION_PATTERN.fullmatch(clean_cover_value(cell.text))
            ),
            None,
        )
        if version_cell is not None:
            result.append((version_cell.ref, version_cell.row, version_cell.col, version_cell.text))
            continue

        inferred_col = date_cell.col - 1
        if inferred_col > output_cell.col:
            result.append((f"{col_to_name(inferred_col)}{date_cell.row}", date_cell.row, inferred_col, ""))

    return result


def replace_excel_drawing_header_values(
    drawing_xml: str,
    old_title: str | None,
    new_title: str | None,
    old_project_title: str | None,
    new_project_title: str | None,
    old_document_numbers: list[str],
    new_document_number: str,
) -> tuple[bytes, int, int, int]:
    text_matches = list(DRAWING_TEXT_PATTERN.finditer(drawing_xml))
    if not text_matches:
        return drawing_xml.encode("utf-8"), 0, 0, 0

    texts = [unescape(match.group("text")) for match in text_matches]
    new_texts: dict[int, str] = {}
    paragraphs = drawing_paragraph_text_runs(drawing_xml, text_matches, texts)
    add_unlabeled_drawing_version_update(new_texts, texts)

    title_count = 0
    project_count = 0
    document_number_count = 0
    old_title_clean = clean_cover_value(old_title or "")
    new_title_clean = clean_cover_value(new_title or "")
    old_project_clean = clean_cover_value(old_project_title or "")
    new_project_clean = clean_cover_value(new_project_title or "")
    for paragraph_index, (run_indexes, combined) in enumerate(paragraphs):
        clean_combined = clean_cover_value(combined)
        if old_title_clean and new_title_clean and old_title_clean != new_title_clean and clean_combined == old_title_clean:
            replace_drawing_paragraph(new_texts, texts, run_indexes, new_title or "")
            title_count += 1
            continue

        if (
            old_project_clean
            and new_project_clean
            and old_project_clean != new_project_clean
            and clean_combined == old_project_clean
        ):
            replace_drawing_paragraph(new_texts, texts, run_indexes, new_project_title or "")
            project_count += 1
            continue

        id_matches = list(OUTPUT_ID_PATTERN.finditer(combined))
        if not id_matches:
            continue

        for match in id_matches:
            if match.group(0) == clean_cover_value(new_document_number):
                continue
            replace_drawing_text_span(new_texts, texts, run_indexes, match.span(), new_document_number)
            document_number_count += 1

        if not new_project_clean:
            continue
        previous = previous_non_empty_paragraph(paragraphs, paragraph_index)
        if previous is None:
            continue
        previous_run_indexes, previous_text = previous
        previous_clean = clean_cover_value(previous_text)
        if (
            previous_clean
            and previous_clean != new_project_clean
            and not any(run_index in new_texts for run_index in previous_run_indexes)
        ):
            replace_drawing_paragraph(new_texts, texts, previous_run_indexes, new_project_title or "")
            project_count += 1

    if not new_texts:
        return drawing_xml.encode("utf-8"), 0, 0, 0

    pieces: list[str] = []
    last = 0
    for index, match in enumerate(text_matches):
        if index not in new_texts:
            continue
        pieces.append(drawing_xml[last:match.start("text")])
        pieces.append(escape(new_texts[index]))
        last = match.end("text")
    pieces.append(drawing_xml[last:])
    return "".join(pieces).encode("utf-8"), title_count, project_count, document_number_count


def add_unlabeled_drawing_version_update(new_texts: dict[int, str], texts: list[str]) -> None:
    combined = "".join(texts)
    date_match = select_drawing_metadata_date(combined)
    if date_match is None:
        return
    version_match = select_drawing_metadata_version(combined, date_match.start())
    if version_match is None:
        return
    replace_drawing_text_span(
        new_texts,
        texts,
        list(range(len(texts))),
        version_match.span(),
        UNLABELED_HEADER_VERSION_VALUE,
    )


def select_drawing_metadata_date(text: str) -> re.Match[str] | None:
    for match in re.finditer(r"\d{4}[-./]\d{1,2}[-./]\d{1,2}", text):
        prefix = text[:match.start()]
        if not OUTPUT_ID_PATTERN.search(prefix[-200:]):
            continue
        if not re.search(r"[vV]?\d+(?:\.\d+)+\s*$", prefix):
            continue
        return match
    return None


def select_drawing_metadata_version(text: str, before_index: int) -> re.Match[str] | None:
    prefix = text[:before_index]
    candidates = [
        match
        for match in re.finditer(r"[vV]?\d+(?:\.\d+)+", prefix)
        if OUTPUT_ID_PATTERN.search(prefix[max(0, match.start() - 200):match.start()])
    ]
    return candidates[-1] if candidates else None


def drawing_paragraph_text_runs(
    drawing_xml: str,
    text_matches: list[re.Match[str]],
    texts: list[str],
) -> list[tuple[list[int], str]]:
    paragraphs: list[tuple[list[int], str]] = []
    for paragraph_match in DRAWING_PARAGRAPH_PATTERN.finditer(drawing_xml):
        run_indexes = [
            index
            for index, text_match in enumerate(text_matches)
            if paragraph_match.start() <= text_match.start() and text_match.end() <= paragraph_match.end()
        ]
        if not run_indexes:
            continue
        combined = "".join(texts[index] for index in run_indexes)
        if clean_cover_value(combined):
            paragraphs.append((run_indexes, combined))
    return paragraphs


def previous_non_empty_paragraph(
    paragraphs: list[tuple[list[int], str]],
    current_index: int,
) -> tuple[list[int], str] | None:
    for index in range(current_index - 1, -1, -1):
        if clean_cover_value(paragraphs[index][1]):
            return paragraphs[index]
    return None


def replace_drawing_paragraph(
    new_texts: dict[int, str],
    texts: list[str],
    run_indexes: list[int],
    replacement: str,
) -> None:
    if not run_indexes:
        return
    combined = "".join(texts[index] for index in run_indexes)
    leading = combined[: len(combined) - len(combined.lstrip())]
    trailing = combined[len(combined.rstrip()):]
    for position, run_index in enumerate(run_indexes):
        new_texts[run_index] = f"{leading}{replacement}{trailing}" if position == 0 else ""


def replace_drawing_text_span(
    new_texts: dict[int, str],
    texts: list[str],
    run_indexes: list[int],
    target_span: tuple[int, int],
    replacement: str,
) -> None:
    target_start, target_end = target_span
    spans: list[tuple[int, int, int]] = []
    cursor = 0
    for run_index in run_indexes:
        start = cursor
        cursor += len(texts[run_index])
        spans.append((run_index, start, cursor))

    matched_spans = [
        (run_index, start, end)
        for run_index, start, end in spans
        if start < target_end and end > target_start
    ]
    for index, (run_index, start, end) in enumerate(matched_spans):
        text = texts[run_index]
        prefix = text[: max(0, target_start - start)] if index == 0 else ""
        suffix = text[max(0, target_end - start):] if index == len(matched_spans) - 1 and target_end < end else ""
        new_texts[run_index] = f"{prefix}{replacement if index == 0 else ''}{suffix}"
