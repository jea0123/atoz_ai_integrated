# 실제 폴더를 돌며 매칭 대상 파일 목록을 수집합니다.
from __future__ import annotations

from pathlib import Path

from .content_identity import read_file_identity
from .folder_policy import FolderPolicy
from .models import ScannedFile


DEFAULT_FILE_SUFFIXES = {
    ".hwp",
    ".hwpx",
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".doc",
    ".docx",
}


def scan_folder(
    folder: Path,
    *,
    recursive: bool = True,
    allowed_suffixes: set[str] | None = None,
    read_contents: bool = True,
    folder_policy: FolderPolicy | None = None,
) -> list[ScannedFile]:
    """폴더 안의 검사 대상 파일 목록을 읽는다."""
    if not folder.exists() or not folder.is_dir():
        raise RuntimeError(f"검사할 폴더가 없습니다: {folder}")

    suffixes = {suffix.lower() for suffix in (allowed_suffixes or DEFAULT_FILE_SUFFIXES)}
    iterator = folder.rglob("*") if recursive else folder.glob("*")
    files: list[ScannedFile] = []
    for path in iterator:
        if not path.is_file() or path.suffix.lower() not in suffixes:
            continue
        if folder_policy and not folder_policy.should_scan(path, folder):
            continue
        identity = read_file_identity(path) if read_contents else None
        files.append(ScannedFile(path, identity))
    return sorted(files, key=lambda item: str(item.path).casefold())
