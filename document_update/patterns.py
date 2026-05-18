# 산출물 식별자와 한글/엑셀 문서 조각을 찾기 위한 공통 정규식을 모아둔다.
from __future__ import annotations

import re


OUTPUT_ID_PATTERN = re.compile(
    r"\b[A-Za-z]{2,10}(?:-[A-Za-z0-9]{1,12}){2,6}\b"
)

ROW_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?tr)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
CELL_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?tc)\b[^>]*>.*?</(?P=tag)>", re.DOTALL)
TEXT_NODE_PATTERN = re.compile(
    r"<(?P<tag>(?:\w+:)?t)\b(?P<attrs>[^>]*?)(?<!/)>(?P<body>.*?)</(?P=tag)>",
    re.DOTALL,
)
RUN_OPEN_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?run)\b(?P<attrs>[^>]*)>")
RUN_SELF_CLOSING_PATTERN = re.compile(r"<(?P<tag>(?:\w+:)?run)\b(?P<attrs>[^>]*)/>")
