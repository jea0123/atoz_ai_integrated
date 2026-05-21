# 대상 한글/엑셀 파일 내부를 열어 문서번호, 프로젝트명, 문서제목을 실제로 치환한다.
# HWPX/Excel 파일 내부의 문서번호, 제목, 프로젝트명을 실제로 교체합니다.
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import shutil
import zipfile
from xml.sax.saxutils import escape, unescape

from .excel_ooxml import EXCEL_DOCUMENT_SUFFIXES, build_updated_excel_cover_sheet
from .hwpx_text import is_hwpx_zip
from .patterns import (
    CELL_PATTERN,
    OUTPUT_ID_PATTERN,
    ROW_PATTERN,
    RUN_OPEN_PATTERN,
    RUN_SELF_CLOSING_PATTERN,
    TEXT_NODE_PATTERN,
)
from .ppt_ooxml import PPT_DOCUMENT_SUFFIXES, write_updated_ppt_document


DOCUMENT_NUMBER_LABEL = "\ubb38\uc11c\ubc88\ud638"


def infer_document_number_from_filename(file_path: Path) -> str:
    """문서번호 셀이 비어 있을 때 대상 파일명에서 문서번호를 보조로 추정한다."""
    match = OUTPUT_ID_PATTERN.search(file_path.stem)
    if match:
        return match.group(0)

    return ""


def unique_scan_values(*values: str) -> list[str]:
    result: list[str] = []
    for value in values:
        value = value.strip()
        if value and value not in result:
            result.append(value)

    return result


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


def replace_document_number_in_section(xml: str, new_document_number: str) -> tuple[str, str]:
    """문서번호 행을 찾아 바로 오른쪽 셀에 새 문서번호를 쓴다."""
    # 문서번호 라벨 바로 오른쪽 셀만 덮어쓴다.
    for row_match in ROW_PATTERN.finditer(xml):
        row_xml = row_match.group(0)
        cells = list(CELL_PATTERN.finditer(row_xml))

        for cell_index, cell_match in enumerate(cells[:-1]):
            label_text = get_cell_text(cell_match.group(0)).strip()
            if label_text != DOCUMENT_NUMBER_LABEL:
                continue

            target_cell_match = cells[cell_index + 1]
            updated_cell, old_document_number = replace_cell_text(
                target_cell_match.group(0),
                new_document_number,
            )
            updated_row = (
                row_xml[:target_cell_match.start()]
                + updated_cell
                + row_xml[target_cell_match.end():]
            )
            updated_xml = xml[:row_match.start()] + updated_row + xml[row_match.end():]
            return updated_xml, old_document_number.strip()

    raise RuntimeError("첫 번째 장에서 문서번호 바로 오른쪽 ID 칸을 찾지 못했습니다.")


def replace_matching_text_nodes(xml: str, old_text: str, new_text: str) -> tuple[str, int]:
    """머리글 같은 곳에 반복된 동일 텍스트 노드를 정확히 일치할 때만 바꾼다."""
    if not old_text or old_text == new_text:
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
        return xml, 0

    pieces.append(xml[last:])
    return "".join(pieces), replace_count


def replace_title_text(xml: str, old_title: str | None, new_title: str | None) -> tuple[str, int]:
    """기존 제목 텍스트 노드를 표준에서 찾은 새 제목으로 바꾼다."""
    if not old_title or not new_title:
        return xml, 0

    if old_title == new_title:
        return xml, 0

    return replace_matching_text_nodes(xml, old_title, new_title)


def backup_path_for(file_path: Path) -> Path:
    return file_path.parent / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_path.suffix}"


def write_updated_document(
    file_path: Path,
    new_document_number: str,
    old_title: str | None = None,
    new_title: str | None = None,
    old_project_title: str | None = None,
    new_project_title: str | None = None,
    output_path: Path | None = None,
) -> tuple[str, Path, int, int, int, Path]:
    """대상 파일 형식에 맞는 수정 함수로 분기한다."""
    suffix = file_path.suffix.lower()

    if is_hwpx_zip(file_path):
        return write_updated_hwpx_document(
            file_path,
            new_document_number,
            old_title,
            new_title,
            old_project_title,
            new_project_title,
            output_path,
        )

    if suffix in EXCEL_DOCUMENT_SUFFIXES:
        return write_updated_excel_document(
            file_path,
            new_document_number,
            old_title,
            new_title,
            old_project_title,
            new_project_title,
            output_path,
        )

    if suffix in PPT_DOCUMENT_SUFFIXES:
        return write_updated_ppt_document(
            file_path,
            new_document_number,
            old_title,
            new_title,
            old_project_title,
            new_project_title,
            output_path,
        )

    raise RuntimeError(
        "지원하지 않는 대상 파일 형식입니다. "
        "대상 파일은 HWP/HWPX, XLSX/XLSM 또는 PPT/PPTX 계열을 사용해주세요."
    )


# 한글 확장 문서는 압축 파일 안에 문서 조각이 들어 있으므로 내부 텍스트 노드를 수정한다.
def write_updated_hwpx_document(
    file_path: Path,
    new_document_number: str,
    old_title: str | None = None,
    new_title: str | None = None,
    old_project_title: str | None = None,
    new_project_title: str | None = None,
    output_path: Path | None = None,
) -> tuple[str, Path, int, int, int, Path]:
    """업로드된 대상 파일을 기준으로 수정된 한글 확장 결과 파일을 만든다."""
    backup_path = backup_path_for(file_path)
    write_path = output_path or file_path.parent / "working_output.hwpx"

    with zipfile.ZipFile(file_path, "r") as source_zip:
        section_xml = source_zip.read("Contents/section0.xml").decode("utf-8", errors="ignore")
        _, old_document_number = replace_document_number_in_section(
            section_xml,
            new_document_number,
        )

    document_numbers_to_scan = unique_scan_values(
        old_document_number,
        infer_document_number_from_filename(file_path),
    )

    if output_path is None:
        # 원본을 덮어쓸 때는 먼저 시각이 들어간 백업을 만든다.
        shutil.copy2(file_path, backup_path)

    title_replace_count = 0
    project_title_replace_count = 0
    matching_document_number_replace_count = 0
    replaced = False

    try:
        with zipfile.ZipFile(file_path, "r") as zin:
            with zipfile.ZipFile(write_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename.lower().endswith(".xml"):
                        xml = data.decode("utf-8", errors="ignore")
                        changed = False

                        # 문서제목, 프로젝트명, 기존 문서번호를 순서대로 바꾼다.
                        xml, count = replace_title_text(xml, old_title, new_title)
                        if count:
                            title_replace_count += count
                            changed = True

                        xml, count = replace_title_text(
                            xml,
                            old_project_title,
                            new_project_title,
                        )
                        if count:
                            project_title_replace_count += count
                            changed = True

                        if item.filename == "Contents/section0.xml":
                            xml, old_document_number = replace_document_number_in_section(
                                xml,
                                new_document_number,
                            )
                            replaced = True
                            changed = True

                        for document_number_to_scan in document_numbers_to_scan:
                            xml, count = replace_matching_text_nodes(
                                xml,
                                document_number_to_scan,
                                new_document_number,
                            )
                            if count:
                                matching_document_number_replace_count += count
                                changed = True

                        if changed:
                            data = xml.encode("utf-8")

                    zout.writestr(item, data)

        if not replaced:
            raise RuntimeError("첫 번째 장에서 문서번호 바로 오른쪽 ID 칸을 찾지 못했습니다.")

        updated_path = write_path
        if output_path is None:
            write_path.replace(file_path)
            updated_path = file_path

        return (
            old_document_number,
            backup_path,
            title_replace_count,
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
    old_title: str | None = None,
    new_title: str | None = None,
    old_project_title: str | None = None,
    new_project_title: str | None = None,
    output_path: Path | None = None,
) -> tuple[str, Path, int, int, int, Path]:
    """엑셀 통합문서의 표지 시트 셀을 수정한다."""
    backup_path = backup_path_for(file_path)
    write_path = output_path or file_path.parent / f"working_output{file_path.suffix}"

    if output_path is None:
        shutil.copy2(file_path, backup_path)

    try:
        with zipfile.ZipFile(file_path, "r") as zin:
            (
                cover_sheet_path,
                updated_cover_sheet,
                old_document_number,
                title_replace_count,
                project_title_replace_count,
                document_number_replace_count,
            ) = build_updated_excel_cover_sheet(
                zin,
                new_document_number,
                old_title,
                new_title,
                old_project_title,
                new_project_title,
            )

            with zipfile.ZipFile(write_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == cover_sheet_path:
                        data = updated_cover_sheet
                    zout.writestr(item, data)

        updated_path = write_path
        if output_path is None:
            write_path.replace(file_path)
            updated_path = file_path

        return (
            old_document_number,
            backup_path,
            title_replace_count,
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
