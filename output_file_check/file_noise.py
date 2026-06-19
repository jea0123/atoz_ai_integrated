from __future__ import annotations

from pathlib import Path


NOISE_FILE_NAMES = {"thumbs.db", "desktop.ini"}


def is_noise_filename(filename: str) -> bool:
    name = Path(filename).name.casefold()
    return (
        name in NOISE_FILE_NAMES
        or name.startswith("~$")
        or name.startswith(".metadata_")
        or "_metadata_" in name
    )


def copytree_ignore_noise(_directory: str, names: list[str]) -> set[str]:
    return {name for name in names if is_noise_filename(name)}


def remove_noise_files(root: Path) -> int:
    if not root.exists():
        return 0

    removed_count = 0
    for path in root.rglob("*"):
        if not path.is_file() or not is_noise_filename(path.name):
            continue
        try:
            path.unlink()
            removed_count += 1
        except OSError:
            pass
    return removed_count
