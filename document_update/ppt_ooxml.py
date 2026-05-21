# PowerPoint 표지 슬라이드에서 프로젝트명/문서명/문서번호를 읽고 텍스트를 치환한다.
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, unescape

from .patterns import OUTPUT_ID_PATTERN, TEXT_NODE_PATTERN


PPT_OOXML_SUFFIXES = {
    ".pptx",
    ".pptm",
    ".potx",
    ".potm",
    ".ppsx",
    ".ppsm",
}
PPT_LEGACY_SUFFIXES = {".ppt"}
PPT_DOCUMENT_SUFFIXES = PPT_OOXML_SUFFIXES | PPT_LEGACY_SUFFIXES

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}

METADATA_LABELS = ("문서번호", "Version", "버전", "개정일자", "작성자")
DOCUMENT_NUMBER_LABEL = "문서번호"
PARAGRAPH_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?p)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
TABLE_ROW_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?tr)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
TABLE_CELL_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?tc)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
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
    "교육자료",
)


@dataclass(frozen=True)
class PptShapeText:
    text: str
    x: int
    y: int
    font_size: int


@dataclass(frozen=True)
class PptCoverIdentity:
    project_title: str = ""
    document_title: str = ""
    document_number: str = ""
    preview_text: str = ""


def backup_path_for(file_path: Path) -> Path:
    return file_path.parent / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{file_path.suffix}"


def read_ppt_cover_identity(file_path: Path) -> PptCoverIdentity:
    """PPTX 첫 번째 슬라이드에서 표지에 보이는 프로젝트명/문서명/문서번호를 읽는다."""
    if file_path.suffix.lower() not in PPT_OOXML_SUFFIXES:
        return PptCoverIdentity(preview_text=file_path.name)

    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            slide_xml = zf.read("ppt/slides/slide1.xml").decode("utf-8", errors="ignore")
    except (KeyError, OSError, zipfile.BadZipFile) as exc:
        raise RuntimeError(f"PPT 첫 번째 슬라이드를 읽지 못했습니다: {file_path.name}") from exc

    shapes = extract_first_slide_shapes(slide_xml)
    lines = [shape.text for shape in shapes if shape.text]
    preview_text = "\n".join(lines)
    return PptCoverIdentity(
        project_title=find_cover_project_title(shapes),
        document_title=find_cover_document_title(shapes),
        document_number=find_cover_document_number(lines),
        preview_text=preview_text,
    )


def extract_first_slide_shapes(slide_xml: str) -> list[PptShapeText]:
    root = ET.fromstring(slide_xml)
    shape_tree = root.find(".//p:spTree", NS)
    if shape_tree is None:
        return []

    shapes: list[PptShapeText] = []
    for element in list(shape_tree):
        text = shape_display_text(element)
        if not text:
            continue
        x, y = element_position(element)
        shapes.append(PptShapeText(text=text, x=x, y=y, font_size=max_font_size(element)))

    return sorted(shapes, key=lambda shape: (shape.y, shape.x))


def shape_display_text(element: ET.Element) -> str:
    paragraphs: list[str] = []
    for paragraph in element.findall(".//a:p", NS):
        pieces = [
            text.text or ""
            for text in paragraph.findall(".//a:t", NS)
            if text.text
        ]
        paragraph_text = clean_cover_text("".join(pieces))
        if paragraph_text:
            paragraphs.append(paragraph_text)
    return clean_cover_text(" ".join(paragraphs))


def element_position(element: ET.Element) -> tuple[int, int]:
    off = element.find(".//a:xfrm/a:off", NS)
    if off is None:
        off = element.find(".//p:xfrm/a:off", NS)
    if off is None:
        return 0, 0
    return int(off.attrib.get("x", "0") or 0), int(off.attrib.get("y", "0") or 0)


def max_font_size(element: ET.Element) -> int:
    sizes: list[int] = []
    for run_properties in element.findall(".//a:rPr", NS):
        raw_size = run_properties.attrib.get("sz", "")
        if raw_size.isdigit():
            sizes.append(int(raw_size))
    return max(sizes, default=0)


def find_cover_project_title(shapes: list[PptShapeText]) -> str:
    title_shapes = cover_title_shapes(shapes)
    document_shape = select_document_title_shape(title_shapes)
    if document_shape is None:
        return ""

    combined_project_title, combined_document_title = split_combined_project_document(document_shape.text)
    if combined_project_title and combined_document_title:
        return combined_project_title

    previous_shapes = [shape for shape in title_shapes if shape.y < document_shape.y]
    if previous_shapes:
        return clean_cover_text(previous_shapes[-1].text)

    index = title_shapes.index(document_shape)
    if index > 0:
        return clean_cover_text(title_shapes[index - 1].text)
    return ""


def find_cover_document_title(shapes: list[PptShapeText]) -> str:
    title_shape = select_document_title_shape(cover_title_shapes(shapes))
    if title_shape is None:
        return ""

    _project_title, document_title = split_combined_project_document(title_shape.text)
    return document_title or clean_cover_text(title_shape.text)


def cover_title_shapes(shapes: list[PptShapeText]) -> list[PptShapeText]:
    metadata_y_values = [shape.y for shape in shapes if has_metadata_label(shape.text)]
    metadata_y = min(metadata_y_values) if metadata_y_values else None
    candidates: list[PptShapeText] = []

    for shape in shapes:
        if metadata_y is not None and shape.y >= metadata_y:
            continue
        if is_cover_noise(shape.text):
            continue
        if is_date_like(shape.text):
            continue
        candidates.append(shape)

    return candidates or [
        shape
        for shape in shapes
        if not is_cover_noise(shape.text) and not has_metadata_label(shape.text)
    ]


def select_document_title_shape(shapes: list[PptShapeText]) -> PptShapeText | None:
    if not shapes:
        return None

    keyword_shapes = [shape for shape in shapes if any(keyword in shape.text for keyword in TITLE_KEYWORDS)]
    pool = keyword_shapes or shapes
    return max(pool, key=lambda shape: (shape.font_size, shape.y))


def split_combined_project_document(value: str) -> tuple[str, str]:
    """'프로젝트명 업무정의서'처럼 한 텍스트 박스에 붙은 표지 제목을 나눈다."""
    text = clean_cover_text(value)
    parts = text.split()
    if len(parts) < 2:
        return "", ""

    last = parts[-1]
    if not any(keyword in last for keyword in TITLE_KEYWORDS):
        return "", ""

    title_parts = [last]
    project_parts = parts[:-1]
    if last == "매뉴얼" and project_parts and project_parts[-1] in {"사용자", "운영자", "관리자"}:
        title_parts.insert(0, project_parts.pop())

    project_title = clean_cover_text(" ".join(project_parts))
    document_title = clean_cover_text(" ".join(title_parts))
    if not project_title or not document_title:
        return "", ""
    return project_title, document_title


def find_cover_document_number(lines: list[str]) -> str:
    joined = " ".join(lines)
    label_match = re.search(r"문서번호\s*[:：]?\s*(?P<value>[A-Za-z0-9][A-Za-z0-9-]+)", joined)
    if label_match:
        return clean_cover_text(label_match.group("value"))

    match = OUTPUT_ID_PATTERN.search(joined)
    return match.group(0) if match else ""


def clean_cover_text(value: str) -> str:
    text = unescape(value)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" -\t\r\n")


def has_metadata_label(value: str) -> bool:
    return any(label in value for label in METADATA_LABELS)


def is_cover_noise(value: str) -> bool:
    text = clean_cover_text(value)
    if not text:
        return True
    if text.startswith("식품의약품안전처"):
        return True
    if text.startswith("(주)") or text.startswith("㈜"):
        return True
    if OUTPUT_ID_PATTERN.fullmatch(text):
        return True
    if has_metadata_label(text):
        return True
    return False


def is_date_like(value: str) -> bool:
    text = clean_cover_text(value)
    compact = re.sub(r"[\s.년월/-]+", "", text)
    return bool(re.fullmatch(r"(?:20)?\d{2}(?:\d{1,2})?", compact))


def replace_matching_text_nodes(xml: str, old_text: str | None, new_text: str | None) -> tuple[str, int]:
    if not old_text or not new_text or old_text == new_text:
        return xml, 0

    escaped_new_text = escape(new_text)
    pieces: list[str] = []
    last = 0
    replace_count = 0

    for match in TEXT_NODE_PATTERN.finditer(xml):
        body = match.group("body")
        if clean_cover_text(body) != old_text:
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

        combined = clean_cover_text("".join(unescape(match.group("body")) for match in text_nodes))
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


def xml_fragment_text(xml: str) -> str:
    return clean_cover_text("".join(unescape(match.group("body")) for match in TEXT_NODE_PATTERN.finditer(xml)))


def replace_cell_text(cell_xml: str, new_text: str) -> tuple[str, str]:
    text_nodes = list(TEXT_NODE_PATTERN.finditer(cell_xml))
    old_text = xml_fragment_text(cell_xml)
    escaped_new_text = escape(new_text)

    if not text_nodes:
        return cell_xml, old_text

    pieces: list[str] = []
    last = 0
    for index, match in enumerate(text_nodes):
        pieces.append(cell_xml[last:match.start("body")])
        pieces.append(escaped_new_text if index == 0 else "")
        last = match.end("body")
    pieces.append(cell_xml[last:])
    return "".join(pieces), old_text


def replace_document_number_cell(xml: str, new_document_number: str) -> tuple[str, str, int]:
    """PPT 표에서 '문서번호' 바로 오른쪽 셀 값을 바꾼다."""
    for row_match in TABLE_ROW_PATTERN.finditer(xml):
        row_xml = row_match.group(0)
        cells = list(TABLE_CELL_PATTERN.finditer(row_xml))

        for cell_index, cell_match in enumerate(cells[:-1]):
            label_text = xml_fragment_text(cell_match.group(0))
            if label_text != DOCUMENT_NUMBER_LABEL:
                continue

            target_cell_match = cells[cell_index + 1]
            updated_cell, old_document_number = replace_cell_text(
                target_cell_match.group(0),
                new_document_number,
            )
            if updated_cell == target_cell_match.group(0):
                return xml, old_document_number, 0

            updated_row = (
                row_xml[:target_cell_match.start()]
                + updated_cell
                + row_xml[target_cell_match.end():]
            )
            updated_xml = xml[:row_match.start()] + updated_row + xml[row_match.end():]
            return updated_xml, old_document_number, 1

    return xml, "", 0


def write_updated_ppt_document(
    file_path: Path,
    new_document_number: str,
    old_title: str | None = None,
    new_title: str | None = None,
    old_project_title: str | None = None,
    new_project_title: str | None = None,
    output_path: Path | None = None,
) -> tuple[str, Path, int, int, int, Path]:
    """PowerPoint OOXML 문서의 첫 표지값을 읽고 동일 텍스트 노드를 치환한다."""
    backup_path = backup_path_for(file_path)
    write_path = output_path or file_path.parent / f"working_output{file_path.suffix}"

    if output_path is None:
        shutil.copy2(file_path, backup_path)

    if file_path.suffix.lower() not in PPT_OOXML_SUFFIXES:
        if file_path.resolve() != write_path.resolve():
            shutil.copy2(file_path, write_path)
        updated_path = write_path
        if output_path is None:
            write_path.replace(file_path)
            updated_path = file_path
        return ("", backup_path, 0, 0, 0, updated_path)

    cover = read_ppt_cover_identity(file_path)
    old_document_number = cover.document_number
    old_title = old_title or cover.document_title
    old_project_title = old_project_title or cover.project_title

    title_replace_count = 0
    project_title_replace_count = 0
    document_number_replace_count = 0

    try:
        with zipfile.ZipFile(file_path, "r") as zin:
            with zipfile.ZipFile(write_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.lower().endswith(".xml"):
                        xml = data.decode("utf-8", errors="ignore")
                        changed = False

                        xml, count = replace_matching_text_nodes(xml, old_title, new_title)
                        if count:
                            title_replace_count += count
                            changed = True

                        xml, count = replace_matching_text_nodes(xml, old_project_title, new_project_title)
                        if count:
                            project_title_replace_count += count
                            changed = True

                        xml, replaced_document_number, count = replace_document_number_cell(xml, new_document_number)
                        if count:
                            if not old_document_number:
                                old_document_number = replaced_document_number
                            document_number_replace_count += count
                            changed = True

                        xml, count = replace_matching_text_nodes(xml, old_document_number, new_document_number)
                        if count:
                            document_number_replace_count += count
                            changed = True

                        if changed:
                            data = xml.encode("utf-8")
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
