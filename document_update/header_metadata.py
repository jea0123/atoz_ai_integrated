from __future__ import annotations

from collections.abc import Callable, Iterable
import re


TextCleaner = Callable[[str], str]
LabelNormalizer = Callable[[str], str]
AuthorValuePredicate = Callable[[str], bool]


def normalized_label_values(
    values: Iterable[str],
    normalize_label: LabelNormalizer,
) -> set[str]:
    return {normalize_label(item) for item in values}


def has_unlabeled_metadata_signal(
    version_value: str,
    date_value: str,
    author_value: str,
    *,
    clean_text: TextCleaner,
) -> bool:
    return bool(clean_text(version_value) or clean_text(date_value) or clean_text(author_value))


def unlabeled_header_metadata_indexes(
    values: list[str],
    *,
    clean_text: TextCleaner,
    normalize_label: LabelNormalizer,
    label_like_values: Iterable[str],
    version_pattern: re.Pattern[str],
    date_pattern: re.Pattern[str],
    author_value_ok: AuthorValuePredicate | None = None,
) -> tuple[int, int, int, int] | None:
    if len(values) < 4:
        return None

    label_like_keys = normalized_label_values(label_like_values, normalize_label)
    for start in range(0, len(values) - 3):
        code_value = clean_text(values[start])
        version_value = clean_text(values[start + 1])
        date_value = clean_text(values[start + 2])
        author_value = clean_text(values[start + 3])
        if (
            _slot_is_clean(
                code_value,
                "code",
                normalize_label=normalize_label,
                label_like_keys=label_like_keys,
                version_pattern=version_pattern,
                date_pattern=date_pattern,
                author_value_ok=author_value_ok,
            )
            and _slot_is_clean(
                version_value,
                "version",
                normalize_label=normalize_label,
                label_like_keys=label_like_keys,
                version_pattern=version_pattern,
                date_pattern=date_pattern,
                author_value_ok=author_value_ok,
            )
            and _slot_is_clean(
                date_value,
                "date",
                normalize_label=normalize_label,
                label_like_keys=label_like_keys,
                version_pattern=version_pattern,
                date_pattern=date_pattern,
                author_value_ok=author_value_ok,
            )
            and _slot_is_clean(
                author_value,
                "author",
                normalize_label=normalize_label,
                label_like_keys=label_like_keys,
                version_pattern=version_pattern,
                date_pattern=date_pattern,
                author_value_ok=author_value_ok,
            )
            and bool(version_value or date_value or author_value)
        ):
            return start, start + 1, start + 2, start + 3
    return None


def unlabeled_header_slot_is_clean(
    value: str,
    slot: str,
    *,
    clean_text: TextCleaner,
    normalize_label: LabelNormalizer,
    label_like_values: Iterable[str],
    version_pattern: re.Pattern[str],
    date_pattern: re.Pattern[str],
    author_value_ok: AuthorValuePredicate | None = None,
) -> bool:
    return _slot_is_clean(
        clean_text(value),
        slot,
        normalize_label=normalize_label,
        label_like_keys=normalized_label_values(label_like_values, normalize_label),
        version_pattern=version_pattern,
        date_pattern=date_pattern,
        author_value_ok=author_value_ok,
    )


def unlabeled_document_code_value_ok(
    value: str,
    *,
    clean_text: TextCleaner,
    normalize_label: LabelNormalizer,
    label_like_values: Iterable[str],
    date_pattern: re.Pattern[str],
) -> bool:
    text = clean_text(value)
    return (
        bool(text)
        and not _is_label_like(
            text,
            normalize_label=normalize_label,
            label_like_keys=normalized_label_values(label_like_values, normalize_label),
        )
        and not date_pattern.fullmatch(text)
    )


def unlabeled_author_value_ok(
    value: str,
    *,
    clean_text: TextCleaner,
    normalize_label: LabelNormalizer,
    label_like_values: Iterable[str],
    version_pattern: re.Pattern[str],
    date_pattern: re.Pattern[str],
) -> bool:
    text = clean_text(value)
    return (
        bool(text)
        and not _is_label_like(
            text,
            normalize_label=normalize_label,
            label_like_keys=normalized_label_values(label_like_values, normalize_label),
        )
        and not date_pattern.fullmatch(text)
        and not version_pattern.fullmatch(text)
    )


def _slot_is_clean(
    text: str,
    slot: str,
    *,
    normalize_label: LabelNormalizer,
    label_like_keys: set[str],
    version_pattern: re.Pattern[str],
    date_pattern: re.Pattern[str],
    author_value_ok: AuthorValuePredicate | None,
) -> bool:
    if not text:
        return True
    if _is_label_like(text, normalize_label=normalize_label, label_like_keys=label_like_keys):
        return False
    if slot == "code":
        return not date_pattern.fullmatch(text)
    if slot == "version":
        return bool(version_pattern.fullmatch(text))
    if slot == "date":
        return bool(date_pattern.fullmatch(text))
    if slot == "author":
        if author_value_ok is not None:
            return author_value_ok(text)
        return not version_pattern.fullmatch(text) and not date_pattern.fullmatch(text)
    return False


def _is_label_like(
    text: str,
    *,
    normalize_label: LabelNormalizer,
    label_like_keys: set[str],
) -> bool:
    return normalize_label(text) in label_like_keys
