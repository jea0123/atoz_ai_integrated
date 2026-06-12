# 웹 multipart 업로드를 파싱하고 임시 작업 폴더에 안전한 파일명으로 저장합니다.
from __future__ import annotations

from email.parser import BytesParser
from email.policy import default
from pathlib import Path
import re

from output_file_check.file_noise import is_noise_filename


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
