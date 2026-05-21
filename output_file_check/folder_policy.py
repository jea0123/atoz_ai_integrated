# 제외 폴더, 투명 폴더, 특정 경로 이하만 검사하는 폴더 정책입니다.
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from .normalization import normalize_for_match


ARCHIVE_FOLDER_NAMES = {"2024", "0.1", "0.0", "구버전", "old"}
REVIEW_FOLDER_NAMES = {"양식", "회신용"}
PATH_REVIEW_KEYWORDS = {"검토표시", "자체테스트"}
ARCHIVE_FILENAME_PREFIXES = ("(2024)", "（2024）")
LOW_VERSION_FILENAME_PATTERN = re.compile(r"(?:^|[_-])v0\.9(?=$|[^0-9])", re.IGNORECASE)


@dataclass(frozen=True)
class FolderPolicy:
    """실제 폴더 중 매핑 대상/제외/보조 폴더를 구분하는 규칙."""

    ignore_folder_names: tuple[str, ...] = ("bak", "backup", "백업", "임시", "temp", "tmp")
    transparent_folder_names: tuple[str, ...] = ("원본",)
    map_only_under: tuple[tuple[str, ...], ...] = ()

    def should_scan(self, path: Path, root_folder: Path) -> bool:
        parts = relative_parent_parts(path, root_folder)
        if is_default_excluded_path(parts, path.name):
            return False

        normalized_parts = {normalize_for_match(part) for part in parts}
        ignored = {normalize_for_match(part) for part in self.ignore_folder_names}
        if normalized_parts & ignored:
            return False

        if not self.map_only_under:
            return True

        return any(has_prefix(parts, prefix) for prefix in self.map_only_under)

    def comparable_path_parts(self, parts: tuple[str, ...]) -> tuple[str, ...]:
        transparent = {normalize_for_match(part) for part in self.transparent_folder_names}
        return tuple(part for part in parts if normalize_for_match(part) not in transparent)

def relative_parent_parts(path: Path, root_folder: Path) -> tuple[str, ...]:
    parent = path.parent
    try:
        relative = parent.resolve().relative_to(root_folder.resolve())
        parts = relative.parts
    except ValueError:
        parts = parent.parts
    return tuple(part for part in parts if part not in {"", "."})


def has_prefix(parts: tuple[str, ...], prefix: tuple[str, ...]) -> bool:
    normalized_parts = [normalize_for_match(part) for part in parts]
    normalized_prefix = [normalize_for_match(part) for part in prefix]
    if len(normalized_parts) < len(normalized_prefix):
        return False
    return normalized_parts[: len(normalized_prefix)] == normalized_prefix


def has_contiguous_path(parts: tuple[str, ...], target: tuple[str, ...]) -> bool:
    normalized_parts = [normalize_for_match(part) for part in parts]
    normalized_target = [normalize_for_match(part) for part in target]
    if not normalized_target or len(normalized_parts) < len(normalized_target):
        return False
    for index in range(len(normalized_parts) - len(normalized_target) + 1):
        if normalized_parts[index:index + len(normalized_target)] == normalized_target:
            return True
    return False


def is_default_excluded_path(parts: tuple[str, ...], filename: str = "") -> bool:
    if has_contiguous_path(parts, ("05.시험", "02.시험단계점검", "02.데이터값진단")):
        return True

    if is_archive_or_review_path(parts, filename):
        return True

    manual_root = ("06.인도", "01.인수인계", "02.매뉴얼작성")
    if has_contiguous_path(parts, manual_root):
        return not any(normalize_for_match(part) == normalize_for_match("운영자매뉴얼") for part in parts)

    return False


def is_archive_or_review_path(parts: tuple[str, ...], filename: str) -> bool:
    normalized_parts = {normalize_for_match(part) for part in parts}
    archive_names = {normalize_for_match(part) for part in ARCHIVE_FOLDER_NAMES}
    review_names = {normalize_for_match(part) for part in REVIEW_FOLDER_NAMES}
    if normalized_parts & (archive_names | review_names):
        return True

    joined_path = normalize_for_match("\\".join(parts))
    if any(normalize_for_match(keyword) in joined_path for keyword in PATH_REVIEW_KEYWORDS):
        return True

    name = filename.strip()
    if any(name.startswith(prefix) for prefix in ARCHIVE_FILENAME_PREFIXES):
        return True

    normalized_name = normalize_for_match(name)
    if "검토표시" in normalized_name or "회신용" in normalized_name:
        return True

    if LOW_VERSION_FILENAME_PATTERN.search(name):
        return True

    return False
