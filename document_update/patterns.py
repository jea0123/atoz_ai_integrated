# 산출물 식별자와 한글/엑셀 문서 조각을 찾기 위한 공통 정규식을 모아둔다.
from __future__ import annotations

import re


OUTPUT_ID_PATTERN_TEXT = r"(?:MFDS-\d{3,6}|[A-Za-z]{2,10}(?:-[A-Za-z0-9]{1,12})*-\d{2})"
NUMBER_OUTPUT_ID_PATTERN_TEXT = r"\d{4,}(?:\([^)]*\))?"
OUTPUT_ID_PATTERN = re.compile(
    rf"\b{OUTPUT_ID_PATTERN_TEXT}\b"
)


def split_output_id_and_name(value: str) -> tuple[str, str]:
    text = str(value or "").strip(" -\t\r\n")
    if "-" not in text:
        return text, ""
    output_id, output_name = text.rsplit("-", 1)
    output_id = output_id.strip(" -\t\r\n")
    output_name = output_name.strip(" -\t\r\n")
    if output_name.isdigit() or not (
        OUTPUT_ID_PATTERN.fullmatch(output_id)
        or re.fullmatch(NUMBER_OUTPUT_ID_PATTERN_TEXT, output_id)
    ):
        return text, ""
    return output_id, output_name

ROW_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?tr)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
CELL_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?tc)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
TEXT_NODE_PATTERN = re.compile(
    r"<(?P<tag>(?:\w+:)?t)\b(?P<attrs>[^>]*?)(?<!/)>(?P<body>.*?)</(?P=tag)>",
    re.DOTALL,
)
RUN_OPEN_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?run)\b(?P<attrs>[^>]*)>")
RUN_SELF_CLOSING_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?run)\b(?P<attrs>[^>]*)/>")
