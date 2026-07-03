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
ARTIFACT_UPLOAD_SUPPORTED_SUFFIXES = {".hwp", ".hwpx", ".docx", ".docm", ".pptx", ".pptm", ".xlsx", ".xlsm", ".xltx", ".xltm"}
PROPOSAL_SUPPORTED_SUFFIXES = ARTIFACT_UPLOAD_SUPPORTED_SUFFIXES | {".pdf"}
CHECK_FOLDER_SUPPORTED_SUFFIXES = ARTIFACT_UPLOAD_SUPPORTED_SUFFIXES | {".xls", ".doc", ".ppt", ".potx", ".potm", ".ppsx", ".ppsm"}
IGNORED_UPLOAD_FOLDER_KEYS = {"bak", "backup", "old"}
PROPOSAL_OOXML_TEXT_SUFFIXES = {".docx", ".docm"}
REQUIREMENT_LIST_MARKER_PATTERN = re.compile(
    r"요구\s*사항\s*(?:목록\s*표?|리스트|명세|정의)|기능\s*요구\s*사항",
    re.IGNORECASE,
)
TOLERANT_REQUIREMENT_ID_PATTERN = re.compile(
    r"(?<![A-Z0-9])S\s*F\s*R\s*[-_]\s*"
    r"(?P<body>(?:[A-Z0-9]+\s*[-_]\s*)*\d+)(?![A-Z0-9])",
    re.IGNORECASE,
)
STRICT_REQUIREMENT_ID_PATTERN = re.compile(r"(?<![A-Z0-9])SFR-(?:[A-Z0-9]+-)*\d+(?![A-Z0-9])", re.IGNORECASE)
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


def normalize_upload_folder_part(value: str) -> str:
    text = unicodedata.normalize("NFKC", value or "")
    text = re.sub(r"\s+", "", text)
    return text.casefold()


def has_ignored_upload_folder(path: Path) -> bool:
    return any(normalize_upload_folder_part(part) in IGNORED_UPLOAD_FOLDER_KEYS for part in path.parts[:-1])


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
        if suffix not in CHECK_FOLDER_SUPPORTED_SUFFIXES:
            continue

        relative_path = safe_relative_upload_path(filename, f"file-{index}{suffix}")
        if has_ignored_upload_folder(relative_path):
            continue
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

        requirement_match_details = extract_requirement_id_match_details_from_proposal_text(text)
        requirement_matches = requirement_match_details["kept"]
        if not requirement_matches:
            saved_paths.append(
                write_requirement_marker(
                    marker_dir,
                    f"ID없음_{index}.txt",
                    f"{filename}: 제안요청서의 요구사항목록표에서 SFR 요구사항 ID를 찾지 못했습니다.",
                )
            )
            continue

        for match_info in requirement_matches:
            requirement_id = str(match_info["requirement_id"])
            if requirement_id in seen_ids:
                continue
            seen_ids.add(requirement_id)
            saved_paths.append(
                write_requirement_marker(
                    marker_dir,
                    f"{requirement_id}.txt",
                    "\n".join(
                        [
                            f"source: {filename}",
                            f"requirement_id: {requirement_id}",
                            f"matched_text: {match_info.get('matched_text', '')}",
                            f"context: {match_info.get('context', '')}",
                        ]
                    ),
                )
            )
        for ignored_index, ignored_match in enumerate(requirement_match_details["ignored"], start=1):
            saved_paths.append(
                write_requirement_marker(
                    marker_dir,
                    f"ID제외_{index}_{ignored_index}.txt",
                    "\n".join(
                        [
                            f"source: {filename}",
                            f"ignored_requirement_id: {ignored_match.get('requirement_id', '')}",
                            f"matched_text: {ignored_match.get('matched_text', '')}",
                            f"context: {ignored_match.get('context', '')}",
                            f"ignore_reason: {ignored_match.get('ignore_reason', '')}",
                        ]
                    ),
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


def extract_requirement_id_match_details_from_proposal_text(text: str) -> dict[str, list[dict[str, str]]]:
    # 요구사항목록표 주변을 우선 보고, 표제가 잘 안 잡히면 문서 전체에서 SFR ID를 찾는다.
    normalized_text = normalize_requirement_text(text)
    section_text = requirement_list_section(normalized_text)
    details = extract_requirement_id_match_details_from_text(section_text) if section_text else empty_requirement_match_details()
    return details if details["kept"] else extract_requirement_id_match_details_from_text(normalized_text)


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
    return tuple(item["requirement_id"] for item in extract_requirement_id_match_details_from_text(text)["kept"])


def extract_requirement_id_match_details_from_text(text: str) -> dict[str, list[dict[str, str]]]:
    seen: set[str] = set()
    found: list[tuple[int, int, str, str]] = []
    candidates: list[dict[str, str]] = []

    for match in STRICT_REQUIREMENT_ID_PATTERN.finditer(text):
        found.append((match.start(), match.end(), match.group(0).upper(), match.group(0)))

    for match in TOLERANT_REQUIREMENT_ID_PATTERN.finditer(text):
        body = re.sub(r"\s*[-_]\s*", "-", match.group("body")).strip("-")
        requirement_id = f"SFR-{body}".upper()
        if not extract_requirement_ids(requirement_id):
            continue
        found.append((match.start(), match.end(), requirement_id, match.group(0)))

    for start, end, requirement_id, matched_text in sorted(found, key=lambda item: item[0]):
        if requirement_id in seen:
            continue
        candidates.append(
            {
                "requirement_id": requirement_id,
                "matched_text": compact_match_text(matched_text),
                "context": match_context(text, start, end),
            }
        )
        seen.add(requirement_id)

    return split_requirement_id_shape_conflicts(candidates)


def empty_requirement_match_details() -> dict[str, list[dict[str, str]]]:
    return {"kept": [], "ignored": []}


def split_requirement_id_shape_conflicts(matches: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    if len(matches) < 2:
        return {"kept": matches, "ignored": []}

    shape_counts: dict[tuple[str, ...], int] = {}
    for item in matches:
        shape = requirement_id_shape(str(item.get("requirement_id") or ""))
        shape_counts[shape] = shape_counts.get(shape, 0) + 1
    dominant_shape, dominant_count = max(shape_counts.items(), key=lambda item: item[1])
    if dominant_count < 2 or list(shape_counts.values()).count(dominant_count) > 1:
        return {"kept": matches, "ignored": []}

    dominant_ids = [
        str(item.get("requirement_id") or "")
        for item in matches
        if requirement_id_shape(str(item.get("requirement_id") or "")) == dominant_shape
    ]
    kept: list[dict[str, str]] = []
    ignored: list[dict[str, str]] = []
    for item in matches:
        requirement_id = str(item.get("requirement_id") or "")
        if requirement_id_shape(requirement_id) == dominant_shape:
            kept.append(item)
            continue
        reason = shape_conflict_reason(requirement_id, dominant_ids)
        if reason:
            ignored.append({**item, "ignore_reason": reason})
        else:
            kept.append(item)
    return {"kept": kept, "ignored": ignored}


def requirement_id_shape(requirement_id: str) -> tuple[str, ...]:
    parts = requirement_id.upper().split("-")
    return tuple(requirement_id_segment_shape(part) for part in parts)


def requirement_id_segment_shape(part: str) -> str:
    if part.isdigit():
        return "N"
    if part.isalpha():
        return "A"
    return "X"


def shape_conflict_reason(requirement_id: str, dominant_ids: list[str]) -> str:
    for dominant_id in dominant_ids:
        if prefix_related_requirement_ids(requirement_id, dominant_id):
            return f"다수 ID 형태와 다르고 {dominant_id}와 접두 관계라 제외"
        if same_last_number(requirement_id, dominant_id):
            return f"다수 ID 형태와 다르고 {dominant_id}와 마지막 숫자가 같아 제외"
    return ""


def prefix_related_requirement_ids(left: str, right: str) -> bool:
    left_key = left.upper()
    right_key = right.upper()
    return left_key.startswith(f"{right_key}-") or right_key.startswith(f"{left_key}-")


def same_last_number(left: str, right: str) -> bool:
    left_number = last_requirement_number(left)
    right_number = last_requirement_number(right)
    return left_number is not None and left_number == right_number


def last_requirement_number(requirement_id: str) -> int | None:
    parts = requirement_id.upper().split("-")
    for part in reversed(parts):
        if part.isdigit():
            return int(part)
    return None


def compact_match_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def match_context(text: str, start: int, end: int, window: int = 80) -> str:
    left = max(0, start - window)
    right = min(len(text), end + window)
    return compact_match_text(text[left:right])


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
