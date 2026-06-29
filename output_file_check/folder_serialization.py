# 매칭 결과를 웹 화면과 CLI가 쓰는 JSON 형태로 바꿉니다.
from __future__ import annotations

from pathlib import Path

from output_file_check.folder_mapping import FolderMappingResult


def serialize_check_result(
    request_id: str,
    standard_file: Path,
    folder_dir: Path,
    mapping: FolderMappingResult,
) -> dict[str, object]:
    # 내부 매칭 결과를 check.js가 바로 그릴 수 있는 응답 JSON으로 바꾼다.
    matched = [item for item in mapping.matches if item.candidates]
    unmatched_reference = [item for item in mapping.matches if not item.candidates]
    matched_file_count = sum(len(item.candidates) for item in matched)

    return {
        "request_id": request_id,
        "standard_file": standard_file.name,
        "standard_project_title": mapping.standard_project_title,
        "artifact_category": mapping.artifact_category,
        "match_mode": mapping.match_mode,
        "scanned_files": len(mapping.files),
        "reference_output_count": len(mapping.outputs),
        "output_count": len(mapping.outputs),
        "path_template_count": len(mapping.path_templates),
        "matched_output_count": len(matched),
        "matched_file_count": matched_file_count,
        "unmatched_reference_output_count": len(unmatched_reference),
        "matches": [
            {
                "output_id": item.output.output_id,
                "output_name": item.output.output_name,
                "folder_name": item.output.folder_name,
                "candidates": [
                    serialize_check_candidate(candidate, folder_dir)
                    for candidate in item.candidates
                ],
            }
            for item in matched
        ],
        "unmatched_reference": [
            {
                "output_id": item.output.output_id,
                "output_name": item.output.output_name,
                "folder_name": item.output.folder_name,
            }
            for item in unmatched_reference
        ],
    }


def serialize_check_candidate(candidate: object, folder_dir: Path) -> dict[str, object]:
    # 후보 하나의 경로, 점수, AI confidence, 표지 정보를 JSON으로 바꾼다.
    try:
        relative_path = str(candidate.file.path.relative_to(folder_dir))
    except ValueError:
        relative_path = str(candidate.file.path)

    identity = candidate.file.identity
    ai_confidence = getattr(candidate, "ai_confidence", None)
    return {
        "path": relative_path,
        "score": candidate.score,
        "ai_confidence": ai_confidence,
        "source": "ai" if ai_confidence is not None else "rule",
        "reason": candidate.reason,
        "identity": {
            "project_title": identity.project_title,
            "document_title": identity.document_title,
            "document_number": identity.document_number,
            "error": identity.error,
        } if identity else None,
    }
