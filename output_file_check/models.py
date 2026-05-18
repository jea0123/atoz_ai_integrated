# 폴더 매칭과 반영 단계가 공유하는 데이터 구조를 모아둔 파일입니다.
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PathTemplate:
    """문서관리표준에서 읽은 산출물의 예상 폴더 경로."""

    output_name: str
    template_path: tuple[str, ...]
    source: str = ""


@dataclass(frozen=True)
class StandardOutput:
    """문서관리표준에서 읽은 산출물 한 건."""

    output_id: str
    output_name: str
    folder_name: str | None = None
    aliases: tuple[str, ...] = ()
    source_line: str = ""


@dataclass(frozen=True)
class FileIdentity:
    """파일 내부에서 읽은 표지/본문 식별 정보."""

    project_title: str = ""
    document_title: str = ""
    preview_text: str = ""
    error: str = ""


@dataclass(frozen=True)
class ScannedFile:
    """검사 대상 폴더에서 찾은 파일 한 건."""

    path: Path
    identity: FileIdentity | None = None

    @property
    def stem(self) -> str:
        return self.path.stem

    @property
    def suffix(self) -> str:
        return self.path.suffix


@dataclass(frozen=True)
class MatchCandidate:
    """산출물과 파일 사이의 매칭 후보."""

    output: StandardOutput
    file: ScannedFile
    score: float
    reason: str
    ai_confidence: float | None = None


@dataclass(frozen=True)
class OutputMatch:
    """산출물별 최종 매칭 결과."""

    output: StandardOutput
    candidates: tuple[MatchCandidate, ...]

    @property
    def candidate(self) -> MatchCandidate | None:
        return self.candidates[0] if self.candidates else None

