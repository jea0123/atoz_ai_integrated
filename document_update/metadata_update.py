from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path
import re
import shutil
import zipfile

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .document_number import get_cell_text, replace_cell_text
from .excel_ooxml import (
    EXCEL_DOCUMENT_SUFFIXES,
    CellInfo,
    col_to_name,
    iter_cells,
    read_shared_strings,
    replace_or_insert_cell_xml,
    workbook_sheets,
)
from .hwpx_text import extract_document_text, is_hwpx_zip
from .patterns import CELL_PATTERN, OUTPUT_ID_PATTERN, ROW_PATTERN, split_output_id_and_name
from .runtime_conversion import prepare_target_file


SUPPORTED_METADATA_SUFFIXES = {".hwp", ".hwpx", *EXCEL_DOCUMENT_SUFFIXES}
IGNORED_FOLDER_NAMES = {"bak", "backup", "font", "KRDS_UIUX", "__pycache__"}
SCHEDULE_SHEET_NAME = "Schedule"
WBS_START_COL = 16
WBS_AUTHOR_COL = 28
WBS_OUTPUT_COL = 29
WBS_TASK_COLS = (5, 6, 7, 8, 9, 10)
DATE_LABELS = {"개정일자"}
AUTHOR_LABELS = {"작성자", "작성 자", "작 성 자"}
REVISION_HEADER_LABELS = {"버전", "개정일자", "작성자"}
APPROVAL_LABELS = {"승인", "승인자"}
LABEL_LIKE_VALUES = {
    "문서번호",
    "문서버전",
    "개정일자",
    "작성자",
    "작성 자",
    "작 성 자",
    "버전",
    "개정사유",
    "개정내역",
    "승인",
    "승인자",
}
VERSION_PATTERN = re.compile(r"^[vV]?\d+(?:\.\d+)*$")
REQUIREMENT_ID_PATTERN = re.compile(r"(?<![A-Z0-9])SFR-[A-Z0-9]+-\d{3}(?!\d)", re.IGNORECASE)
HEADER_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?header)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
DATE_VALUE_PATTERN = re.compile(r"^\d{4}[-./]\d{1,2}[-./]\d{1,2}$")


@dataclass(frozen=True)
class WbsMetadata:
    output_name: str
    author: str
    revision_date: str
    wbs: str
    task: str
    row: int
    requirement_id: str = ""


@dataclass
class DocumentMetadata:
    author: str = ""
    revision_date: str = ""
    revision_author: str = ""
    revision_history_date: str = ""


@dataclass
class MetadataTarget:
    path: Path
    relative_path: str
    status: str
    message: str
    output_name: str = ""
    author: str = ""
    revision_date: str = ""
    current: DocumentMetadata = field(default_factory=DocumentMetadata)
    candidates: list[WbsMetadata] = field(default_factory=list)


@dataclass
class MetadataWriteResult:
    status: str
    old_path: Path
    new_path: Path | None = None
    backup_path: Path | None = None
    converted_to_hwpx: bool = False
    cover_update_count: int = 0
    revision_history_update_count: int = 0
    error: str = ""


def normalize_key(value: str) -> str:
    return re.sub(r"[\s_\-(){}\[\]·.,/]+", "", value or "").casefold()


def normalize_label(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def clean_text(value: object) -> str:
    return str(value or "").strip()


def split_output_names(value: object) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    names: list[str] = []
    seen: set[str] = set()
    for part in re.split(r"[,;\n]+", text):
        raw = part.strip()
        _output_id, output_name = split_output_id_and_name(raw)
        for name in (raw, output_name):
            key = normalize_key(name)
            if not key or key in seen:
                continue
            names.append(name)
            seen.add(key)
    return names


def format_revision_date(value: object) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return datetime.combine(value, time()).strftime("%Y-%m-%d")

    text = clean_text(value)
    if not text:
        return ""

    for pattern in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text[:19], pattern).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return text[:10]


def read_wbs_metadata(wbs_path: Path) -> list[WbsMetadata]:
    workbook = load_workbook(wbs_path, read_only=True, data_only=True, keep_vba=True)
    if SCHEDULE_SHEET_NAME not in workbook.sheetnames:
        raise RuntimeError("WBS 파일에서 Schedule 시트를 찾지 못했습니다.")

    sheet = workbook[SCHEDULE_SHEET_NAME]
    records: list[WbsMetadata] = []
    current_requirement_id = ""
    for row_index, row in enumerate(sheet.iter_rows(min_row=4, values_only=True), start=4):
        row_requirement_id = extract_requirement_id_from_values(row)
        if row_requirement_id:
            current_requirement_id = row_requirement_id

        author = clean_text(row[WBS_AUTHOR_COL - 1] if len(row) >= WBS_AUTHOR_COL else "")
        revision_date = format_revision_date(row[WBS_START_COL - 1] if len(row) >= WBS_START_COL else "")
        output_names = split_output_names(row[WBS_OUTPUT_COL - 1] if len(row) >= WBS_OUTPUT_COL else "")
        if not author or not revision_date or not output_names:
            continue

        wbs = clean_text(row[3] if len(row) >= 4 else "")
        task = next(
            (clean_text(row[column - 1]) for column in reversed(WBS_TASK_COLS) if len(row) >= column and clean_text(row[column - 1])),
            "",
        )
        for output_name in output_names:
            records.append(
                WbsMetadata(
                    output_name=output_name,
                    author=author,
                    revision_date=revision_date,
                    wbs=wbs,
                    task=task,
                    row=row_index,
                    requirement_id=current_requirement_id,
                )
            )
    return records


def extract_requirement_id_from_values(values: object) -> str:
    if isinstance(values, (str, Path)):
        text = str(values)
    else:
        try:
            text = " ".join(clean_text(value) for value in values if value is not None)
        except TypeError:
            text = clean_text(values)

    match = REQUIREMENT_ID_PATTERN.search(text)
    return match.group(0).upper() if match else ""


def build_wbs_index(records: list[WbsMetadata]) -> dict[str, list[WbsMetadata]]:
    index: dict[str, list[WbsMetadata]] = {}
    for record in records:
        index.setdefault(normalize_key(record.output_name), []).append(record)
    return index


def should_scan_document(path: Path) -> bool:
    if path.suffix.lower() not in SUPPORTED_METADATA_SUFFIXES:
        return False
    if any(part in IGNORED_FOLDER_NAMES for part in path.parts):
        return False
    if "WBS" in path.name.upper():
        return False
    return True


def collect_metadata_documents(root: Path) -> list[Path]:
    return sorted(
        [path for path in root.rglob("*") if path.is_file() and should_scan_document(path)],
        key=lambda item: str(item).casefold(),
    )


def records_matching_path(path: Path, records: list[WbsMetadata]) -> list[WbsMetadata]:
    haystack = normalize_key(" ".join(path.parts[-5:]))
    matches = [record for record in records if normalize_key(record.output_name) and normalize_key(record.output_name) in haystack]
    if not matches:
        return []

    best_length = max(len(normalize_key(record.output_name)) for record in matches)
    output_matches = [record for record in matches if len(normalize_key(record.output_name)) == best_length]

    requirement_id = extract_requirement_id_from_values(path.name)
    if requirement_id:
        return [
            record
            for record in output_matches
            if record.requirement_id.casefold() == requirement_id.casefold()
        ]

    return output_matches


def choose_record(
    path: Path,
    candidates: list[WbsMetadata],
    current: DocumentMetadata | None = None,
) -> tuple[WbsMetadata | None, str, str]:
    if not candidates:
        return None, "unmatched", "WBS 산출물명과 파일/폴더명이 매칭되지 않았습니다."

    if not extract_requirement_id_from_values(path.name):
        return earliest_record(candidates), "matched", ""

    path_text = normalize_key(str(path))
    author_matches = [record for record in candidates if normalize_key(record.author) in path_text]
    if len({(item.author, item.output_name) for item in author_matches}) == 1:
        return earliest_record(author_matches), "matched", ""

    current_author = normalize_key(current.author if current else "")
    if current_author:
        current_author_matches = [record for record in candidates if normalize_key(record.author) == current_author]
        if len({(item.author, item.output_name) for item in current_author_matches}) == 1:
            return earliest_record(current_author_matches), "matched", ""

    distinct = {(item.author, item.revision_date, item.output_name) for item in candidates}
    if len(distinct) == 1:
        return candidates[0], "matched", ""

    if len({(item.author, item.output_name) for item in candidates}) == 1:
        return earliest_record(candidates), "matched", ""

    return None, "ambiguous", "같은 산출물명에 서로 다른 WBS 담당/시작일 후보가 있습니다."


def earliest_record(records: list[WbsMetadata]) -> WbsMetadata:
    return sorted(records, key=lambda item: (item.revision_date, item.row))[0]


def inspect_document_metadata(path: Path) -> DocumentMetadata:
    suffix = path.suffix.lower()
    try:
        if suffix in EXCEL_DOCUMENT_SUFFIXES:
            return inspect_excel_metadata(path)
        if is_hwpx_zip(path):
            return inspect_hwpx_metadata(path)
        return inspect_text_metadata(path)
    except Exception:
        return DocumentMetadata()


def inspect_text_metadata(path: Path) -> DocumentMetadata:
    text = extract_document_text(path)[:6000]
    metadata = DocumentMetadata()
    author_match = re.search(r"<\s*(?:작성자|작\s*성\s*자)\s*>\s*<\s*([^<>]+)\s*>", text)
    date_match = re.search(r"<\s*개정일자\s*>\s*<\s*([^<>]+)\s*>", text)
    if author_match:
        metadata.author = author_match.group(1).strip()
    if date_match:
        metadata.revision_date = date_match.group(1).strip()
    return metadata


def inspect_hwpx_metadata(path: Path) -> DocumentMetadata:
    metadata = DocumentMetadata()
    with zipfile.ZipFile(path, "r") as zf:
        for name in zf.namelist():
            if not name.lower().endswith(".xml"):
                continue
            xml = zf.read(name).decode("utf-8", errors="ignore")
            if not metadata.revision_date:
                metadata.revision_date = find_label_value_in_xml(xml, DATE_LABELS)
            if not metadata.author:
                metadata.author = find_label_value_in_xml(xml, AUTHOR_LABELS)
            header_date, header_author = find_header_metadata_values_in_xml(xml)
            if header_date and not metadata.revision_date:
                metadata.revision_date = header_date
            if header_author and not metadata.author:
                metadata.author = header_author
            revision_date, revision_author = find_revision_history_values_in_xml(xml)
            if revision_date and not metadata.revision_history_date:
                metadata.revision_history_date = revision_date
            if revision_author and not metadata.revision_author:
                metadata.revision_author = revision_author
    return metadata


def find_label_value_in_xml(xml: str, labels: set[str]) -> str:
    normalized_labels = {normalize_label(label) for label in labels}
    for row_match in ROW_PATTERN.finditer(xml):
        cells = list(CELL_PATTERN.finditer(row_match.group(0)))
        if len(cells) > 4:
            continue
        for cell_index, cell_match in enumerate(cells[:-1]):
            label = normalize_label(get_cell_text(cell_match.group(0)).strip())
            if label not in normalized_labels:
                continue
            value = get_cell_text(cells[cell_index + 1].group(0)).strip()
            if value and normalize_label(value) not in {normalize_label(item) for item in LABEL_LIKE_VALUES}:
                return value
    return ""


def find_header_metadata_values_in_xml(xml: str) -> tuple[str, str]:
    for header_match in HEADER_PATTERN.finditer(xml):
        header_xml = header_match.group(0)
        for row_match in ROW_PATTERN.finditer(header_xml):
            cells = list(CELL_PATTERN.finditer(row_match.group(0)))
            if len(cells) < 4:
                continue
            values = [get_cell_text(cell.group(0)).strip() for cell in cells]
            if not looks_like_unlabeled_header_metadata_row(values):
                continue
            return values[2], values[3]
    return "", ""


def looks_like_unlabeled_header_metadata_row(values: list[str]) -> bool:
    if len(values) < 4:
        return False
    return (
        bool(OUTPUT_ID_PATTERN.search(values[0]))
        and bool(VERSION_PATTERN.fullmatch(values[1]))
        and bool(DATE_VALUE_PATTERN.fullmatch(values[2].strip()))
        and normalize_label(values[3]) not in {normalize_label(item) for item in LABEL_LIKE_VALUES}
    )


def find_revision_history_values_in_xml(xml: str) -> tuple[str, str]:
    for _row_match, cells, header_map in iter_revision_history_data_rows(xml):
        date_idx = header_map.get("date")
        author_idx = header_map.get("author")
        if date_idx is None or author_idx is None:
            continue
        return get_cell_text(cells[date_idx].group(0)).strip(), get_cell_text(cells[author_idx].group(0)).strip()
    return "", ""


def inspect_excel_metadata(path: Path) -> DocumentMetadata:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_vba=path.suffix.lower() in {".xlsm", ".xltm"})
    metadata = DocumentMetadata()
    for sheet in workbook.worksheets:
        if not metadata.revision_date:
            metadata.revision_date = find_label_value_in_sheet(sheet, DATE_LABELS)
        if not metadata.author:
            metadata.author = find_label_value_in_sheet(sheet, AUTHOR_LABELS)
    for sheet in workbook.worksheets:
        if normalize_key(sheet.title).startswith(normalize_key("개정이력")):
            metadata.revision_history_date, metadata.revision_author = find_revision_history_values_in_sheet(sheet)
            break
    return metadata


def find_label_value_in_sheet(sheet: Worksheet, labels: set[str]) -> str:
    normalized_labels = {normalize_label(label) for label in labels}
    for row in sheet.iter_rows():
        for index, cell in enumerate(row[:-1]):
            if normalize_label(clean_text(cell.value)) not in normalized_labels:
                continue
            target = row[index + 1]
            value = clean_text(target.value)
            if value and normalize_label(value) not in {normalize_label(item) for item in LABEL_LIKE_VALUES}:
                return value
    return ""


def find_revision_history_values_in_sheet(sheet: Worksheet) -> tuple[str, str]:
    for row in sheet.iter_rows():
        labels = [normalize_label(clean_text(cell.value)) for cell in row]
        if not REVISION_HEADER_LABELS.issubset(set(labels)):
            continue
        date_idx = labels.index("개정일자")
        author_idx = labels.index("작성자")
        for data_row in sheet.iter_rows(min_row=row[0].row + 1, max_col=max(date_idx, author_idx) + 1):
            version = clean_text(data_row[0].value)
            if VERSION_PATTERN.fullmatch(version):
                return clean_text(data_row[date_idx].value), clean_text(data_row[author_idx].value)
    return "", ""


def build_metadata_targets(folder_root: Path, wbs_records: list[WbsMetadata]) -> list[MetadataTarget]:
    targets: list[MetadataTarget] = []
    for path in collect_metadata_documents(folder_root):
        relative_path = str(path.relative_to(folder_root))
        candidates = records_matching_path(path, wbs_records)
        current = inspect_document_metadata(path)
        record, status, message = choose_record(path, candidates, current)
        targets.append(
            MetadataTarget(
                path=path,
                relative_path=relative_path,
                status=status,
                message=message,
                output_name=record.output_name if record else (candidates[0].output_name if candidates else ""),
                author=record.author if record else "",
                revision_date=record.revision_date if record else "",
                current=current,
                candidates=candidates,
            )
        )
    return targets


def update_metadata_in_document(
    file_path: Path,
    author: str,
    revision_date: str,
    approval_author: str,
    temp_dir: Path,
) -> MetadataWriteResult:
    backup_path = backup_path_for(file_path)
    shutil.copy2(file_path, backup_path)
    try:
        target_file, converted_to_hwpx = prepare_target_file(file_path, temp_dir)
        output_path = target_file
        if is_hwpx_zip(target_file):
            result = write_updated_hwpx_metadata(target_file, author, revision_date, approval_author)
        elif target_file.suffix.lower() in EXCEL_DOCUMENT_SUFFIXES:
            result = write_updated_excel_metadata(target_file, author, revision_date, approval_author)
        else:
            raise RuntimeError("지원하지 않는 문서 형식입니다.")

        if converted_to_hwpx:
            output_path = file_path.with_suffix(".hwpx")
            shutil.move(str(target_file), str(output_path))
            if output_path != file_path and file_path.exists():
                file_path.unlink()

        return MetadataWriteResult(
            status="updated",
            old_path=file_path,
            new_path=output_path,
            backup_path=backup_path,
            converted_to_hwpx=converted_to_hwpx,
            cover_update_count=result[0],
            revision_history_update_count=result[1],
        )
    except Exception as exc:
        return MetadataWriteResult(
            status="error",
            old_path=file_path,
            backup_path=backup_path,
            error=str(exc),
        )


def backup_path_for(file_path: Path) -> Path:
    backup_dir = file_path.parent / "bak"
    backup_dir.mkdir(exist_ok=True)
    candidate = backup_dir / file_path.name
    if not candidate.exists():
        return candidate
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return backup_dir / f"{file_path.stem}_{timestamp}{file_path.suffix}"


def write_updated_excel_metadata(path: Path, author: str, revision_date: str, approval_author: str) -> tuple[int, int]:
    cover_count = 0
    revision_count = 0
    temp_path = path.with_name(f".metadata_{path.name}")
    try:
        with zipfile.ZipFile(path, "r") as zin:
            shared_strings = read_shared_strings(zin)
            sheet_map = {sheet.path: sheet.name for sheet in workbook_sheets(zin)}
            updates_by_sheet: dict[str, tuple[bytes, int, int]] = {}

            for sheet_path, sheet_name in sheet_map.items():
                original_xml = zin.read(sheet_path).decode("utf-8", errors="ignore")
                is_revision_history_sheet = normalize_key(sheet_name).startswith(normalize_key("개정이력"))
                label_count = 0
                updated_xml = original_xml
                if not is_revision_history_sheet:
                    updated_xml, label_count = update_excel_label_cells_xml(
                        original_xml,
                        shared_strings,
                        {
                            **{label: revision_date for label in DATE_LABELS},
                            **{label: author for label in AUTHOR_LABELS},
                        },
                    )
                sheet_revision_count = 0
                if is_revision_history_sheet:
                    updated_xml, sheet_revision_count = update_excel_revision_history_xml(
                        updated_xml,
                        shared_strings,
                        revision_date,
                        author,
                        approval_author,
                    )
                if label_count or sheet_revision_count:
                    updates_by_sheet[sheet_path] = (
                        updated_xml.encode("utf-8"),
                        label_count,
                        sheet_revision_count,
                    )
                    cover_count += label_count
                    revision_count += sheet_revision_count

            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename in updates_by_sheet:
                        data = updates_by_sheet[item.filename][0]
                    zout.writestr(item, data)
        temp_path.replace(path)
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise
    return cover_count, revision_count


def update_excel_label_cells_xml(
    sheet_xml: str,
    shared_strings: list[str],
    values_by_label: dict[str, str],
) -> tuple[str, int]:
    import xml.etree.ElementTree as ET

    root = ET.fromstring(sheet_xml)
    cells = iter_cells(root, shared_strings)
    cells_by_position = {(cell.row, cell.col): cell for cell in cells}
    normalized_values = {normalize_label(label): value for label, value in values_by_label.items()}
    label_values = {normalize_label(item) for item in LABEL_LIKE_VALUES}
    updates: dict[str, tuple[int, int, str]] = {}

    for info in cells:
        label = normalize_label(clean_text(info.text))
        if label not in normalized_values:
            continue
        target_col = info.col + 1
        target = cells_by_position.get((info.row, target_col))
        if target and normalize_label(clean_text(target.text)) in label_values:
            continue
        target_ref = f"{col_to_name(target_col)}{info.row}"
        updates[target_ref] = (info.row, target_col, normalized_values[label])

    updated_xml = sheet_xml
    for cell_ref, (row, col, value) in updates.items():
        updated_xml = replace_or_insert_cell_xml(updated_xml, cell_ref, row, col, value)
    return updated_xml, len(updates)


def update_excel_revision_history_xml(
    sheet_xml: str,
    shared_strings: list[str],
    revision_date: str,
    author: str,
    approval_author: str,
) -> tuple[str, int]:
    import xml.etree.ElementTree as ET

    root = ET.fromstring(sheet_xml)
    cells = iter_cells(root, shared_strings)
    rows: dict[int, list[CellInfo]] = {}
    for cell in cells:
        rows.setdefault(cell.row, []).append(cell)

    for row_index in sorted(rows):
        row_cells = sorted(rows[row_index], key=lambda item: item.col)
        labels = [normalize_label(clean_text(cell.text)) for cell in row_cells]
        header_by_key = revision_header_map(labels)
        if not header_by_key:
            continue
        header_by_col = {
            key: row_cells[index].col
            for key, index in header_by_key.items()
        }
        first_data_row_index: int | None = None
        updated_xml = sheet_xml
        count = 0
        for data_row_index in sorted(index for index in rows if index > row_index):
            data_cells = {cell.col: cell for cell in rows[data_row_index]}
            version_cell = data_cells.get(header_by_col["version"])
            version = clean_text(version_cell.text if version_cell else "")
            if VERSION_PATTERN.fullmatch(version):
                if first_data_row_index is None:
                    first_data_row_index = data_row_index
                    updates = {
                        "date": revision_date,
                        "author": author,
                        "approval": approval_author,
                    }
                    for key, value in updates.items():
                        col = header_by_col.get(key)
                        if col is None:
                            continue
                        updated_xml = replace_or_insert_cell_xml(
                            updated_xml,
                            f"{col_to_name(col)}{data_row_index}",
                            data_row_index,
                            col,
                            value,
                        )
                        count += 1
                    continue

                min_col = min(header_by_col.values())
                max_col = max(header_by_col.values())
                for cell in sorted(data_cells.values(), key=lambda item: item.col):
                    if min_col <= cell.col <= max_col and clean_text(cell.text):
                        updated_xml = replace_or_insert_cell_xml(
                            updated_xml,
                            cell.ref,
                            data_row_index,
                            cell.col,
                            "",
                        )
                        count += 1
        return updated_xml, count
    return sheet_xml, 0


def revision_header_map(labels: list[str]) -> dict[str, int]:
    normalized_approvals = {normalize_label(label) for label in APPROVAL_LABELS}
    result: dict[str, int] = {}
    for index, label in enumerate(labels):
        if label == normalize_label("버전"):
            result["version"] = index
        elif label == normalize_label("개정일자"):
            result["date"] = index
        elif label == normalize_label("작성자"):
            result["author"] = index
        elif label in normalized_approvals:
            result["approval"] = index
    required = {"version", "date", "author"}
    return result if required.issubset(result) else {}


def update_label_right_cells(sheet: Worksheet, labels: set[str], new_value: object) -> int:
    normalized_labels = {normalize_label(label) for label in labels}
    count = 0
    for row in sheet.iter_rows():
        for index, cell in enumerate(row[:-1]):
            if normalize_label(clean_text(cell.value)) not in normalized_labels:
                continue
            target = row[index + 1]
            if normalize_label(clean_text(target.value)) in {normalize_label(item) for item in LABEL_LIKE_VALUES}:
                continue
            target.value = new_value
            if isinstance(new_value, date):
                target.number_format = "yyyy-mm-dd"
            count += 1
    return count


def update_revision_history_sheet(sheet: Worksheet, revision_date: date, author: str) -> int:
    for row in sheet.iter_rows():
        labels = [normalize_label(clean_text(cell.value)) for cell in row]
        if not REVISION_HEADER_LABELS.issubset(set(labels)):
            continue
        date_idx = labels.index("개정일자")
        author_idx = labels.index("작성자")
        for data_row in sheet.iter_rows(min_row=row[0].row + 1, max_col=max(date_idx, author_idx) + 1):
            version = clean_text(data_row[0].value)
            if VERSION_PATTERN.fullmatch(version):
                data_row[date_idx].value = revision_date
                data_row[date_idx].number_format = "yyyy-mm-dd"
                data_row[author_idx].value = author
                return 2
        return 0
    return 0


def write_updated_hwpx_metadata(path: Path, author: str, revision_date: str, approval_author: str) -> tuple[int, int]:
    temp_path = path.with_name(f".metadata_{path.name}")
    cover_count = 0
    revision_count = 0
    try:
        with zipfile.ZipFile(path, "r") as zin:
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.lower().endswith(".xml"):
                        xml = data.decode("utf-8", errors="ignore")
                        xml, count = update_label_right_rows_in_xml(xml, DATE_LABELS, revision_date)
                        cover_count += count
                        xml, count = update_label_right_rows_in_xml(xml, AUTHOR_LABELS, author)
                        cover_count += count
                        xml, count = update_unlabeled_header_metadata_xml(xml, revision_date, author)
                        cover_count += count
                        xml, count = update_revision_history_xml(xml, revision_date, author, approval_author)
                        revision_count += count
                        data = xml.encode("utf-8")
                    zout.writestr(item, data)
        temp_path.replace(path)
    except Exception:
        if temp_path.exists():
            temp_path.unlink()
        raise
    return cover_count, revision_count


def update_label_right_rows_in_xml(xml: str, labels: set[str], new_text: str) -> tuple[str, int]:
    normalized_labels = {normalize_label(label) for label in labels}
    pieces: list[str] = []
    last = 0
    count = 0

    for row_match in ROW_PATTERN.finditer(xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))
        if not cells or len(cells) > 4:
            continue
        updated_row = row_xml
        offset = 0
        changed = False
        for cell_index, cell_match in enumerate(cells[:-1]):
            label = normalize_label(get_cell_text(cell_match.group(0)).strip())
            if label not in normalized_labels:
                continue
            target_cell = cells[cell_index + 1]
            target_text = normalize_label(get_cell_text(target_cell.group(0)).strip())
            if target_text in {normalize_label(item) for item in LABEL_LIKE_VALUES}:
                continue
            updated_cell, _old_value = replace_cell_text(target_cell.group(0), new_text)
            start = target_cell.start() + offset
            end = target_cell.end() + offset
            updated_row = updated_row[:start] + updated_cell + updated_row[end:]
            offset += len(updated_cell) - len(target_cell.group(0))
            changed = True
            count += 1
        if changed:
            pieces.append(xml[last:row_match.start()])
            pieces.append(updated_row)
            last = row_match.end()

    if not pieces:
        return xml, 0
    pieces.append(xml[last:])
    return "".join(pieces), count


def update_unlabeled_header_metadata_xml(xml: str, revision_date: str, author: str) -> tuple[str, int]:
    pieces: list[str] = []
    last = 0
    count = 0

    for header_match in HEADER_PATTERN.finditer(xml):
        header_xml = header_match.group(0)
        updated_header, header_count = update_unlabeled_header_metadata_block(
            header_xml,
            revision_date,
            author,
        )
        if not header_count:
            continue
        pieces.append(xml[last:header_match.start()])
        pieces.append(updated_header)
        last = header_match.end()
        count += header_count

    if not pieces:
        return xml, 0
    pieces.append(xml[last:])
    return "".join(pieces), count


def update_unlabeled_header_metadata_block(
    header_xml: str,
    revision_date: str,
    author: str,
) -> tuple[str, int]:
    for row_match in ROW_PATTERN.finditer(header_xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))
        if len(cells) < 4:
            continue

        values = [get_cell_text(cell.group(0)).strip() for cell in cells]
        if not looks_like_unlabeled_header_metadata_row(values):
            continue

        header_revision_date = format_date_like_existing(revision_date, values[2])
        updates = [(2, header_revision_date), (3, author)]
        updated_row = row_xml
        offset = 0
        count = 0
        for cell_index, new_text in updates:
            cell = cells[cell_index]
            old_text = get_cell_text(cell.group(0)).strip()
            if old_text == new_text:
                continue
            updated_cell, _old_value = replace_cell_text(cell.group(0), new_text)
            start = cell.start() + offset
            end = cell.end() + offset
            updated_row = updated_row[:start] + updated_cell + updated_row[end:]
            offset += len(updated_cell) - len(cell.group(0))
            count += 1
        if not count:
            return header_xml, 0
        return header_xml[:row_match.start()] + updated_row + header_xml[row_match.end():], count
    return header_xml, 0


def format_date_like_existing(revision_date: str, existing_date: str) -> str:
    normalized = revision_date.strip()
    if "." in existing_date:
        return normalized.replace("-", ".").replace("/", ".")
    if "/" in existing_date:
        return normalized.replace("-", "/").replace(".", "/")
    return normalized.replace(".", "-").replace("/", "-")


def iter_revision_history_data_rows(xml: str):
    rows = list(ROW_PATTERN.finditer(xml))
    for header_index, row_match in enumerate(rows):
        cells = list(CELL_PATTERN.finditer(row_match.group(0)))
        labels = [normalize_label(get_cell_text(cell.group(0)).strip()) for cell in cells]
        header_map = revision_header_map(labels)
        if not header_map:
            continue
        for data_row_match in rows[header_index + 1:]:
            data_cells = list(CELL_PATTERN.finditer(data_row_match.group(0)))
            if len(data_cells) <= max(header_map.values()):
                continue
            version = get_cell_text(data_cells[header_map["version"]].group(0)).strip()
            if VERSION_PATTERN.fullmatch(version):
                yield data_row_match, data_cells, header_map
                break


def update_revision_history_xml(xml: str, revision_date: str, author: str, approval_author: str) -> tuple[str, int]:
    rows = list(ROW_PATTERN.finditer(xml))
    for header_index, row_match in enumerate(rows):
        cells = list(CELL_PATTERN.finditer(row_match.group(0)))
        labels = [normalize_label(get_cell_text(cell.group(0)).strip()) for cell in cells]
        header_map = revision_header_map(labels)
        if not header_map:
            continue

        pieces: list[str] = []
        last = 0
        count = 0
        found_first_data_row = False
        for data_row_match in rows[header_index + 1:]:
            data_cells = list(CELL_PATTERN.finditer(data_row_match.group(0)))
            if len(data_cells) <= max(header_map.values()):
                continue
            version = get_cell_text(data_cells[header_map["version"]].group(0)).strip()
            if not VERSION_PATTERN.fullmatch(version):
                continue

            if found_first_data_row:
                updated_row, row_count = clear_revision_history_data_row(data_row_match.group(0), data_cells)
            else:
                updates = [
                    (header_map["date"], revision_date),
                    (header_map["author"], author),
                ]
                if "approval" in header_map:
                    updates.append((header_map["approval"], approval_author))
                updated_row, row_count = update_revision_history_first_row(
                    data_row_match.group(0),
                    data_cells,
                    updates,
                )
                found_first_data_row = True

            if row_count:
                pieces.append(xml[last:data_row_match.start()])
                pieces.append(updated_row)
                last = data_row_match.end()
                count += row_count

        if not count:
            return xml, 0
        pieces.append(xml[last:])
        return "".join(pieces), count
    return xml, 0


def update_revision_history_first_row(
    row_xml: str,
    cells: list[re.Match[str]],
    updates: list[tuple[int, str]],
) -> tuple[str, int]:
    offset = 0
    updated_row = row_xml
    count = 0
    for cell_index, new_text in updates:
        cell = cells[cell_index]
        updated_cell, _old_value = replace_cell_text(cell.group(0), new_text)
        start = cell.start() + offset
        end = cell.end() + offset
        updated_row = updated_row[:start] + updated_cell + updated_row[end:]
        offset += len(updated_cell) - len(cell.group(0))
        count += 1
    return updated_row, count


def clear_revision_history_data_row(row_xml: str, cells: list[re.Match[str]]) -> tuple[str, int]:
    offset = 0
    updated_row = row_xml
    count = 0
    for cell in cells:
        if not get_cell_text(cell.group(0)).strip():
            continue
        updated_cell, _old_value = replace_cell_text(cell.group(0), "")
        start = cell.start() + offset
        end = cell.end() + offset
        updated_row = updated_row[:start] + updated_cell + updated_row[end:]
        offset += len(updated_cell) - len(cell.group(0))
        count += 1
    return updated_row, count
