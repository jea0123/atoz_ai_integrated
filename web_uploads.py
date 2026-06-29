# 웹 multipart 업로드를 파싱하고 임시 작업 폴더에 안전한 파일명으로 저장합니다.
from __future__ import annotations

from email.parser import BytesParser
from email.policy import default
from pathlib import Path
import re
import unicodedata

from document_update.hwpx_text import extract_document_text, extract_text_from_ooxml, is_zip_container
from output_file_check.file_noise import is_noise_filename
from output_file_check.requirement_generation import extract_requirement_ids


PROPOSAL_FILE_FIELDS = ("proposal_files", "proposal_file", "rfp_files", "rfp_file")
PROPOSAL_SUPPORTED_SUFFIXES = {
    ".pdf",
    ".hwp",
    ".hwpx",
    ".docx",
    ".docm",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
}
PROPOSAL_OOXML_TEXT_SUFFIXES = {".docx", ".docm"}
REQUIREMENT_LIST_MARKER_PATTERN = re.compile(
    r"요구\s*사항\s*(?:목록\s*표?|리스트|명세|정의)|기능\s*요구\s*사항",
    re.IGNORECASE,
)
TOLERANT_REQUIREMENT_ID_PATTERN = re.compile(
    r"(?<![A-Z0-9])S\s*F\s*R\s*[-_\s]*"
    r"(?P<body>[A-Z0-9]+(?:[-_\s]+[A-Z0-9]+)*[-_\s]+\d+)(?![A-Z0-9])",
    re.IGNORECASE,
)
REQUIREMENT_LIST_WINDOW_CHARS = 100_000


def parse_multipart_items(content_type: str, body: bytes) -> tuple[dict[str, str], dict[str, list[tuple[str, bytes]]]]:
    # multipart/form-data 본문을 일반 필드와 파일 목록으로 분리한다.
    """웹 프레임워크 없이 브라우저 폼 데이터를 파싱한다."""
    message = BytesParser(policy=default).parsebytes(
        b"Content-Type: " + content_type.encode("utf-8") + b"\r\n\r\n" + body
    )

    fields: dict[str, str] = {}
    files: dict[str, list[tuple[str, bytes]]] = {}

    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue

        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""

        if filename:
            files.setdefault(name, []).append((filename, payload))
        else:
            fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="ignore")

    return fields, files


def safe_upload_filename(filename: str, field_name: str, suffix: str) -> str:
    # 업로드 파일명이 비었거나 경로 문자를 포함해도 안전한 파일명으로 바꾼다.
    """제목 추론에 쓸 원래 파일명은 보존하되 경로 조작 문자는 막는다."""
    candidate = Path(filename).name.strip()
    if not candidate or candidate in {".", ".."}:
        candidate = f"{field_name}{suffix}"

    candidate = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", candidate)
    if not Path(candidate).suffix:
        candidate = f"{candidate}{suffix}"
    return candidate


def safe_relative_upload_path(filename: str, fallback_name: str) -> Path:
    # webkitdirectory 상대 경로에서 위험한 절대/상위 경로 이동을 제거한다.
    """폴더 업로드의 상대 경로는 보존하되 상위 경로 탈출 문자는 제거한다."""
    normalized = filename.replace("\\", "/").strip("/")
    parts = [
        re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", part.strip())
        for part in normalized.split("/")
        if part.strip() and part not in {".", ".."}
    ]
    if not parts:
        parts = [fallback_name]
    return Path(*parts)


def save_check_uploads(
    temp_dir: Path,
    file_items: dict[str, list[tuple[str, bytes]]],
    *,
    fallback_folder: Path | None = None,
) -> tuple[Path, Path]:
    # 폴더 검사 화면의 표준 PDF와 여러 폴더 파일을 임시 구조로 저장한다.
    """검사 화면에서 올린 문서관리표준 PDF와 폴더 파일들을 임시 작업 폴더에 저장한다."""
    standard_items = file_items.get("standard_file") or []
    if not standard_items or not standard_items[0][1]:
        raise ValueError("문서관리표준 PDF를 선택하세요.")

    standard_name, standard_payload = standard_items[0]
    if Path(standard_name).suffix.lower() != ".pdf":
        raise ValueError("문서관리표준은 PDF만 허용됩니다.")

    standard_path = temp_dir / safe_upload_filename(standard_name, "standard_file", ".pdf")
    standard_path.write_bytes(standard_payload)

    folder_dir = temp_dir / "u"
    folder_dir.mkdir()
    folder_items = file_items.get("folder_files") or []
    if not folder_items and fallback_folder is not None:
        if not fallback_folder.exists() or not fallback_folder.is_dir():
            raise ValueError(f"기본 검사 폴더를 찾을 수 없습니다: {fallback_folder}")
        return standard_path, fallback_folder

    saved_count = 0
    for index, (filename, payload) in enumerate(folder_items, start=1):
        if not payload:
            continue
        if is_noise_filename(filename):
            continue
        suffix = Path(filename).suffix.lower()

        relative_path = safe_relative_upload_path(filename, f"file-{index}{suffix}")
        target_path = folder_dir / relative_path
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(payload)
        saved_count += 1

    if saved_count == 0:
        raise ValueError("업로드된 폴더에서 저장할 파일을 찾지 못했습니다.")

    return standard_path, folder_dir


def save_requirement_uploads(
    temp_dir: Path,
    file_items: dict[str, list[tuple[str, bytes]]],
) -> list[Path]:
    # 요구사항별 자동 생성에 사용할 파일들을 임시 폴더에 저장한다.
    """요구사항 파일 input으로 올라온 파일들을 상대 경로 없이 안전하게 저장한다."""
    requirement_items = file_items.get("requirement_files") or []
    if not requirement_items:
        return []

    requirement_dir = temp_dir / "requirements"
    requirement_dir.mkdir(exist_ok=True)

    saved_paths: list[Path] = []
    for index, (filename, payload) in enumerate(requirement_items, start=1):
        if not payload:
            continue
        if is_noise_filename(filename):
            continue
        suffix = Path(filename).suffix.lower()
        safe_name = safe_upload_filename(filename, "requirement_file", suffix or ".bin")
        target_path = requirement_dir / safe_name
        if target_path.exists():
            target_path = requirement_dir / f"{target_path.stem}_{index}{target_path.suffix}"
        target_path.write_bytes(payload)
        saved_paths.append(target_path)

    return saved_paths


def save_proposal_requirement_uploads(
    temp_dir: Path,
    file_items: dict[str, list[tuple[str, bytes]]],
) -> list[Path]:
    # 제안요청서 본문에서 SFR 요구사항 ID를 읽어 기존 요구사항 생성기가 쓰는 소스 파일로 만든다.
    """제안요청서/RFP 파일 안의 요구사항목록표에서 SFR ID를 추출한다."""
    proposal_items = proposal_upload_items(file_items)
    if not proposal_items:
        return []

    source_dir = temp_dir / "proposal_sources"
    marker_dir = temp_dir / "proposal_requirement_ids"
    source_dir.mkdir(exist_ok=True)
    marker_dir.mkdir(exist_ok=True)

    saved_paths: list[Path] = []
    seen_ids: set[str] = set()
    for index, (filename, payload) in enumerate(proposal_items, start=1):
        if not payload or is_noise_filename(filename):
            continue

        suffix = Path(filename).suffix.lower()
        safe_name = safe_upload_filename(filename, "proposal_file", suffix or ".bin")
        source_path = source_dir / safe_name
        if source_path.exists():
            source_path = source_dir / f"{source_path.stem}_{index}{source_path.suffix}"
        source_path.write_bytes(payload)

        if suffix not in PROPOSAL_SUPPORTED_SUFFIXES:
            saved_paths.append(
                write_requirement_marker(
                    marker_dir,
                    f"ID없음_{index}.txt",
                    f"{filename}: 지원하지 않는 제안요청서 형식입니다: {suffix or '확장자 없음'}",
                )
            )
            continue

        try:
            text = extract_proposal_text(source_path)
        except Exception as exc:
            saved_paths.append(
                write_requirement_marker(
                    marker_dir,
                    f"ID없음_{index}.txt",
                    f"{filename}: 제안요청서 텍스트를 읽지 못했습니다: {exc}",
                )
            )
            continue

        requirement_ids = extract_requirement_ids_from_proposal_text(text)
        if not requirement_ids:
            saved_paths.append(
                write_requirement_marker(
                    marker_dir,
                    f"ID없음_{index}.txt",
                    f"{filename}: 제안요청서의 요구사항목록표에서 SFR 요구사항 ID를 찾지 못했습니다.",
                )
            )
            continue

        for requirement_id in requirement_ids:
            if requirement_id in seen_ids:
                continue
            seen_ids.add(requirement_id)
            saved_paths.append(
                write_requirement_marker(
                    marker_dir,
                    f"{requirement_id}.txt",
                    f"source: {filename}\nrequirement_id: {requirement_id}",
                )
            )

    return saved_paths


def proposal_upload_items(file_items: dict[str, list[tuple[str, bytes]]]) -> list[tuple[str, bytes]]:
    items: list[tuple[str, bytes]] = []
    for field_name in PROPOSAL_FILE_FIELDS:
        items.extend(file_items.get(field_name) or [])
    return items


def extract_proposal_text(path: Path) -> str:
    if path.suffix.lower() in PROPOSAL_OOXML_TEXT_SUFFIXES:
        if not is_zip_container(path):
            raise RuntimeError(f"Office XML 형식이 아닙니다: {path.name}")
        return extract_text_from_ooxml(path)
    return extract_document_text(path)


def extract_requirement_ids_from_proposal_text(text: str) -> tuple[str, ...]:
    # 요구사항목록표 주변을 우선 보고, 표제가 잘 안 잡히면 문서 전체에서 SFR ID를 찾는다.
    normalized_text = normalize_requirement_text(text)
    section_text = requirement_list_section(normalized_text)
    ids = extract_requirement_ids_from_text(section_text) if section_text else ()
    return ids or extract_requirement_ids_from_text(normalized_text)


def normalize_requirement_text(text: str) -> str:
    value = unicodedata.normalize("NFKC", text or "")
    return (
        value
        .replace("－", "-")
        .replace("–", "-")
        .replace("—", "-")
        .replace("―", "-")
    )


def requirement_list_section(text: str) -> str:
    matches = list(REQUIREMENT_LIST_MARKER_PATTERN.finditer(text))
    if not matches:
        return ""
    start = matches[0].start()
    return text[start:start + REQUIREMENT_LIST_WINDOW_CHARS]


def extract_requirement_ids_from_text(text: str) -> tuple[str, ...]:
    seen: set[str] = set()
    found: list[tuple[int, str]] = []
    result: list[str] = []

    for match in re.finditer(r"(?<![A-Z0-9])SFR-(?:[A-Z0-9]+-)*\d+(?![A-Z0-9])", text, re.IGNORECASE):
        found.append((match.start(), match.group(0).upper()))

    for match in TOLERANT_REQUIREMENT_ID_PATTERN.finditer(text):
        body = re.sub(r"[^A-Za-z0-9]+", "-", match.group("body")).strip("-")
        requirement_id = f"SFR-{body}".upper()
        if not extract_requirement_ids(requirement_id):
            continue
        found.append((match.start(), requirement_id))

    for _position, requirement_id in sorted(found, key=lambda item: item[0]):
        if requirement_id in seen:
            continue
        result.append(requirement_id)
        seen.add(requirement_id)

    return tuple(result)


def write_requirement_marker(marker_dir: Path, filename: str, text: str) -> Path:
    target_path = marker_dir / safe_upload_filename(filename, "requirement_source", ".txt")
    if target_path.exists():
        for index in range(2, 1000):
            candidate = marker_dir / f"{target_path.stem}_{index}{target_path.suffix}"
            if not candidate.exists():
                target_path = candidate
                break
    target_path.write_text(text, encoding="utf-8")
    return target_path
