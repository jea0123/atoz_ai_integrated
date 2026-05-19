# check.html 결과 폴더를 기준으로 QA 산출물을 생성하고 기존 파일을 bak로 보관합니다.
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import shutil
from uuid import uuid4

from openpyxl import load_workbook

from app_runtime import TEMP_DIR, log_event, remove_runtime_path
from document_update.hwpx_text import extract_document_text
from qa_generation.generate_tc import generate_test_cases
from qa_generation.generate_ts import generate_test_scenarios
from web_uploads import safe_relative_upload_path


IGNORED_FOLDER_NAMES = {"bak", "backup", "백업", "원본"}
DESIGN_DOCUMENT_SUFFIXES = {".hwp", ".hwpx", ".pdf"}
TC_TEMPLATE_SUFFIXES = {".hwpx"}
TS_TEMPLATE_SUFFIXES = {".xlsx"}
QA_SOURCE_SUFFIXES = DESIGN_DOCUMENT_SUFFIXES | TC_TEMPLATE_SUFFIXES | TS_TEMPLATE_SUFFIXES
REQ_ID_PATTERN = re.compile(r"SFR-[A-Z0-9]+(?:-[A-Z0-9]+)*", re.IGNORECASE)
TC_KEYWORDS = (
    "단위시험케이스",
    "단위시험 케이스",
    "단위테스트",
    "단위 테스트",
    "unittestcase",
    "unit test case",
)
TC_EXCLUDE_KEYWORDS = ("통합시험", "통합테스트", "시나리오", "결과서", "인수인계")


class QaFolderMatchingError(ValueError):
    def __init__(self, message: str, payload: dict[str, object]):
        super().__init__(message)
        self.payload = payload


def run_folder_qa_pipeline(
        dump_root: Path,
        *,
        model_name: str,
        ollama_url: str,
        scenario_form_path: Path,
        ui_design_items: list[tuple[str, bytes]] | None = None,
        ui_design_root: Path | None = None,
        qa_source_items: list[tuple[str, bytes]] | None = None,
        qa_source_root: Path | None = None,
        tc_source_root: Path | None = None,
        ts_source_root: Path | None = None,
        request_id: str | None = None,
) -> dict[str, object]:
    # check.html이 만든 결과 폴더 안에서 QA 입력물을 찾아 생성 결과를 같은 위치에 배치한다.
    request_id = request_id or uuid4().hex[:8]
    dump_root = Path(dump_root).expanduser().resolve()
    if not dump_root.exists() or not dump_root.is_dir():
        raise ValueError(f"결과 폴더를 찾지 못했습니다: {dump_root}")

    temp_dir = TEMP_DIR / f"qa-folder-{request_id}"
    tc_output_dir = temp_dir / "tc-output"
    ts_output_dir = temp_dir / "ts-output"
    temp_dir.mkdir(parents=True, exist_ok=True)

    log_event("qa.folder.start", request_id=request_id, dump_root=str(dump_root))

    try:
        ui_design_paths = collect_design_documents(ui_design_root)
        ui_design_paths.extend(save_uploaded_design_documents(temp_dir, ui_design_items or []))
        qa_source_paths = collect_source_documents(qa_source_root)
        qa_source_paths.extend(save_uploaded_source_documents(temp_dir, qa_source_items or []))
        tc_source_paths = collect_documents(tc_source_root, TC_TEMPLATE_SUFFIXES, "단위시험 폴더")
        ts_source_paths = collect_documents(ts_source_root, TS_TEMPLATE_SUFFIXES, "통합시험 폴더")
        selection = select_qa_source_files(
            dump_root,
            ui_design_paths,
            qa_source_paths=qa_source_paths,
            tc_source_paths=tc_source_paths,
            ts_source_paths=ts_source_paths,
        )
        requirement_items = build_requirement_work_items(selection)
        if not requirement_items:
            payload = build_matching_failure_payload(
                request_id=request_id,
                dump_root=dump_root,
                selection=selection,
                message="요구사항 ID 기준으로 함께 처리할 화면설계서, 단위시험케이스, 통합시험시나리오를 찾지 못했습니다.",
            )
            log_event(
                "qa.folder.matching_empty",
                request_id=request_id,
                dump_root=str(dump_root),
                role_counts=payload.get("role_counts"),
            )
            raise QaFolderMatchingError(str(payload["error"]), payload)

        placed_files: list[dict[str, object]] = []
        processed_items: list[dict[str, object]] = []
        total_tc_count = 0
        total_ts_count = 0

        for item in requirement_items:
            requirement_id = str(item["requirement_id"])
            req_temp_dir = temp_dir / safe_requirement_dirname(requirement_id)
            item_tc_output_dir = req_temp_dir / "tc-output"
            item_ts_output_dir = req_temp_dir / "ts-output"
            try:
                ui_design = Path(str(item["ui_design_path"]))
                tc_template = Path(str(item["tc_template_path"]))
                ts_template = Path(str(item["ts_template_path"]))

                tc_payload = generate_test_cases(
                    pdf_path=ui_design,
                    model_name=model_name,
                    ollama_url=ollama_url,
                    output_dir=item_tc_output_dir,
                    template_path=tc_template,
                )
                if not tc_payload.get("ok"):
                    raise RuntimeError(str(tc_payload.get("error") or "단위시험 케이스 생성에 실패했습니다."))

                tc_xlsx = find_generated_file(tc_payload, suffix=".xlsx", kind="xlsx")
                if tc_xlsx is None:
                    raise RuntimeError("통합시험 시나리오 생성에 사용할 단위시험 케이스 XLSX가 생성되지 않았습니다.")

                tc_hwpx = find_generated_file(tc_payload, suffix=".hwpx", kind="hwpx")
                if tc_hwpx is None:
                    raise RuntimeError("교체할 단위시험케이스 HWPX가 생성되지 않았습니다.")

                ts_payload = generate_test_scenarios(
                    template_xlsx_path=ts_template,
                    tc_xlsx_path=tc_xlsx,
                    ui_pdf_path=ui_design,
                    output_dir=item_ts_output_dir,
                    form_path=scenario_form_path,
                )
                if not ts_payload.get("ok"):
                    raise RuntimeError(str(ts_payload.get("error") or "통합시험 시나리오 생성에 실패했습니다."))

                ts_xlsx = find_generated_file(ts_payload, suffix=ts_template.suffix, kind="xlsx")
                if ts_xlsx is None:
                    raise RuntimeError("교체할 통합시험시나리오 XLSX가 생성되지 않았습니다.")

                placed_for_requirement = [
                    place_generated_file(tc_hwpx, tc_template, "tc_hwpx", "단위시험케이스", requirement_id),
                    place_generated_file(ts_xlsx, ts_template, "ts_xlsx", "통합시험시나리오", requirement_id),
                ]
                placed_files.extend(placed_for_requirement)
                tc_count = int(tc_payload.get("count") or 0)
                ts_count = int(ts_payload.get("count") or 0)
                total_tc_count += tc_count
                total_ts_count += ts_count
                processed_items.append({
                    **item,
                    "status": "updated",
                    "tc_count": tc_count,
                    "ts_count": ts_count,
                    "placed_files": placed_for_requirement,
                })
            except Exception as exc:
                processed_items.append({
                    **item,
                    "status": "error",
                    "error": str(exc),
                })

        failed_items = [item for item in processed_items if item.get("status") == "error"]
        payload = {
            "ok": bool(placed_files),
            "request_id": request_id,
            "dump_root": str(dump_root),
            "tc_count": total_tc_count,
            "ts_count": total_ts_count,
            "requirement_count": len(requirement_items),
            "processed_requirement_count": len(requirement_items) - len(failed_items),
            "failed_requirement_count": len(failed_items),
            "source_files": serialize_selection(selection),
            "requirement_items": processed_items,
            "missing_requirements": build_missing_requirement_report(selection),
            "placed_files": placed_files,
        }
        log_event(
            "qa.folder.done",
            request_id=request_id,
            dump_root=str(dump_root),
            requirements=len(requirement_items),
            processed=payload["processed_requirement_count"],
            failed=len(failed_items),
            tc_count=payload["tc_count"],
            ts_count=payload["ts_count"],
            placed=[item["path"] for item in placed_files],
        )
        return payload
    except Exception as exc:
        log_event("qa.folder.error", request_id=request_id, dump_root=str(dump_root), error=str(exc))
        raise
    finally:
        remove_runtime_path(temp_dir)


def select_qa_source_files(
        dump_root: Path,
        ui_design_paths: list[Path],
        *,
        qa_source_paths: list[Path] | None = None,
        tc_source_paths: list[Path] | None = None,
        ts_source_paths: list[Path] | None = None,
) -> dict[str, dict[str, object]]:
    # QA 대상 문서를 요구사항 ID별로 찾는다. 설계서는 업로드 파일을 우선 사용하고, 없으면 결과 폴더에서 보조 탐색한다.
    files = list(iter_candidate_files(dump_root))
    qa_source_paths = list(qa_source_paths or [])
    tc_source_paths = list(tc_source_paths or [])
    ts_source_paths = list(ts_source_paths or [])
    qa_source_ui_files = filter_role_files(
        dump_root,
        qa_source_paths,
        suffixes=DESIGN_DOCUMENT_SUFFIXES,
        keywords=("사용자인터페이스설계서", "사용자 인터페이스 설계서", "화면설계서", "화면정의서", "ui설계서"),
        exclude_keywords=("단위시험", "통합시험", "시나리오", "결과서"),
    )
    qa_source_tc_files = filter_role_files(
        dump_root,
        qa_source_paths,
        suffixes=TC_TEMPLATE_SUFFIXES,
        keywords=TC_KEYWORDS,
        exclude_keywords=TC_EXCLUDE_KEYWORDS,
    )
    qa_source_ts_files = filter_role_files(
        dump_root,
        qa_source_paths,
        suffixes=TS_TEMPLATE_SUFFIXES,
        keywords=(
            "통합시험시나리오",
            "통합시험 시나리오",
            "통합시험결과서",
            "통합시험 결과서",
            "통합테스트",
            "통합 테스트",
            "integrationtestscenario",
            "integration test scenario",
        ),
        exclude_keywords=("단위시험", "단위테스트", "케이스", "인수인계"),
    )
    fallback_ui_files = [
        path for path in files
        if path.suffix.lower() in DESIGN_DOCUMENT_SUFFIXES
        and requirement_ids_from_path(path)
        and score_keywords(
            searchable_text(dump_root, path),
            ("사용자인터페이스설계서", "사용자 인터페이스 설계서", "화면설계서", "화면정의서", "ui설계서"),
        ) > 0
    ]
    return {
        "ui_design": {
            "label": "사용자인터페이스설계서",
            "by_requirement": index_design_files_by_requirement(unique_paths(
                ui_design_paths + qa_source_ui_files + ([] if ui_design_paths or qa_source_ui_files else fallback_ui_files)
            )),
        },
        "tc_template": {
            "label": "단위시험케이스",
            "by_requirement": index_artifact_files_by_requirement(
                dump_root,
                unique_paths(files + qa_source_tc_files + tc_source_paths),
                suffixes=TC_TEMPLATE_SUFFIXES,
                keywords=TC_KEYWORDS,
                exclude_keywords=TC_EXCLUDE_KEYWORDS,
            ),
        },
        "ts_template": {
            "label": "통합시험시나리오",
            "by_requirement": index_artifact_files_by_requirement(
                dump_root,
                unique_paths(files + qa_source_ts_files + ts_source_paths),
                suffixes=TS_TEMPLATE_SUFFIXES,
                keywords=(
                    "통합시험시나리오",
                    "통합시험 시나리오",
                    "통합시험결과서",
                    "통합시험 결과서",
                    "통합테스트",
                    "통합 테스트",
                    "integrationtestscenario",
                    "integration test scenario",
                ),
                exclude_keywords=("단위시험", "단위테스트", "케이스", "인수인계"),
            ),
        },
    }


def iter_candidate_files(root: Path):
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if any(part.casefold() in IGNORED_FOLDER_NAMES for part in path.relative_to(root).parts[:-1]):
            continue
        yield path


def unique_paths(paths: list[Path]) -> list[Path]:
    result: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        resolved = str(Path(path).resolve()).casefold()
        if resolved in seen:
            continue
        seen.add(resolved)
        result.append(Path(path))
    return result


def filter_role_files(
        root: Path,
        files: list[Path],
        *,
        suffixes: set[str],
        keywords: tuple[str, ...],
        exclude_keywords: tuple[str, ...] = (),
) -> list[Path]:
    result: list[Path] = []
    for path in files:
        if path.suffix.lower() not in suffixes:
            continue
        text = searchable_text(root, path)
        if score_keywords(text, keywords) - score_keywords(text, exclude_keywords) <= 0:
            continue
        result.append(path)
    return result


def collect_documents(root: Path | None, suffixes: set[str], label: str) -> list[Path]:
    # 사용자가 지정한 로컬 폴더에서 허용 확장자 파일을 모두 모은다.
    if root is None:
        return []

    root = Path(root).expanduser()
    if not str(root).strip():
        return []

    root = root.resolve()
    if not root.exists() or not root.is_dir():
        raise ValueError(f"{label}를 찾지 못했습니다: {root}")

    return [
        path
        for path in root.rglob("*")
        if path.is_file()
        and path.suffix.lower() in suffixes
        and not any(part.casefold() in IGNORED_FOLDER_NAMES for part in path.relative_to(root).parts[:-1])
    ]


def collect_design_documents(root: Path | None) -> list[Path]:
    # 사용자가 지정한 화면설계서 폴더에서 HWP/HWPX/PDF를 모두 모은다.
    return collect_documents(root, DESIGN_DOCUMENT_SUFFIXES, "화면설계서 폴더")


def collect_source_documents(root: Path | None) -> list[Path]:
    # QA 원천 폴더에서 화면설계서/단위시험케이스/통합시험 파일 후보를 모두 모은다.
    return collect_documents(root, QA_SOURCE_SUFFIXES, "QA 원천 폴더")


def save_uploaded_design_documents(temp_dir: Path, items: list[tuple[str, bytes]]) -> list[Path]:
    # qa.html에서 직접 올린 화면/사용자인터페이스 설계서를 임시 폴더에 저장한다.
    design_dir = temp_dir / "ui-designs"
    saved: list[Path] = []
    for index, (filename, payload) in enumerate(items, start=1):
        if not payload:
            continue
        suffix = Path(filename).suffix.lower()
        if suffix not in DESIGN_DOCUMENT_SUFFIXES:
            continue

        relative_path = safe_relative_upload_path(filename, f"ui-design-{index}{suffix}")
        target_path = design_dir / relative_path
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(payload)
        saved.append(target_path)

    return saved


def save_uploaded_source_documents(temp_dir: Path, items: list[tuple[str, bytes]]) -> list[Path]:
    # qa.html에서 선택한 QA 원천 폴더를 임시 폴더에 상대 경로 그대로 저장한다.
    source_dir = temp_dir / "qa-sources"
    saved: list[Path] = []
    for index, (filename, payload) in enumerate(items, start=1):
        if not payload:
            continue
        suffix = Path(filename).suffix.lower()
        if suffix not in QA_SOURCE_SUFFIXES:
            continue

        relative_path = safe_relative_upload_path(filename, f"qa-source-{index}{suffix}")
        target_path = source_dir / relative_path
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(payload)
        saved.append(target_path)

    return saved


def index_design_files_by_requirement(files: list[Path]) -> dict[str, dict[str, object]]:
    # 업로드된 설계서 파일을 SFR 요구사항 ID별로 묶는다.
    result: dict[str, dict[str, object]] = {}
    for path in files:
        if path.suffix.lower() not in DESIGN_DOCUMENT_SUFFIXES:
            continue

        requirement_ids = requirement_ids_from_path(path)
        if not requirement_ids:
            requirement_ids = requirement_ids_from_document_text(path)

        for requirement_id in requirement_ids:
            score = 1000 + score_keywords(searchable_text(path.parent, path), ("사용자인터페이스설계서", "화면설계서", "화면정의서"))
            upsert_requirement_file(result, requirement_id, path, score)

    return result


def index_artifact_files_by_requirement(
        root: Path,
        files: list[Path],
        *,
        suffixes: set[str],
        keywords: tuple[str, ...],
        exclude_keywords: tuple[str, ...] = (),
) -> dict[str, dict[str, object]]:
    # check 결과 폴더의 QA 산출물 양식을 SFR 요구사항 ID별로 묶는다.
    result: dict[str, dict[str, object]] = {}
    for path in files:
        if path.suffix.lower() not in suffixes:
            continue

        text = searchable_text(root, path)
        requirement_ids = requirement_ids_from_path(path)
        if not requirement_ids:
            requirement_ids = requirement_ids_from_artifact_content(path)
        if not requirement_ids:
            continue

        score = score_keywords(text, keywords) - score_keywords(text, exclude_keywords)
        if score <= 0:
            continue

        for requirement_id in requirement_ids:
            upsert_requirement_file(result, requirement_id, path, score)

    return result


def upsert_requirement_file(result: dict[str, dict[str, object]], requirement_id: str, path: Path, score: int) -> None:
    existing = result.get(requirement_id)
    candidate = {"path": str(path), "score": score}
    if existing is None:
        result[requirement_id] = {
            "path": str(path),
            "score": score,
            "candidates": [candidate],
        }
        return

    candidates = list(existing.get("candidates") or [])
    candidates.append(candidate)
    candidates = sorted(
        candidates,
        key=lambda item: (int(item.get("score") or 0), -len(str(item.get("path") or "")), str(item.get("path") or "").casefold()),
        reverse=True,
    )
    best = candidates[0]
    existing["path"] = best["path"]
    existing["score"] = best["score"]
    existing["candidates"] = candidates[:5]


def requirement_ids_from_path(path: Path) -> list[str]:
    return requirement_ids_from_text(path.as_posix())


def requirement_ids_from_text(value: str) -> list[str]:
    return sorted({
        match.group(0).upper()
        for match in REQ_ID_PATTERN.finditer(value)
    })


def requirement_ids_from_document_text(path: Path) -> list[str]:
    try:
        text = extract_document_text(path)
    except Exception as exc:
        log_event("qa.folder.design_text_error", path=str(path), error=str(exc))
        return []
    return requirement_ids_from_text(text)


def requirement_ids_from_artifact_content(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    if suffix in {".hwp", ".hwpx", ".pdf"}:
        return requirement_ids_from_document_text(path)
    if suffix == ".xlsx":
        return requirement_ids_from_spreadsheet(path)
    return []


def requirement_ids_from_spreadsheet(path: Path) -> list[str]:
    found: set[str] = set()
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        log_event("qa.folder.spreadsheet_text_error", path=str(path), error=str(exc))
        return []

    try:
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    found.update(requirement_ids_from_text(str(cell.value)))
    except Exception as exc:
        log_event("qa.folder.spreadsheet_scan_error", path=str(path), error=str(exc))
    finally:
        workbook.close()

    return sorted(found)


def build_requirement_work_items(selection: dict[str, dict[str, object]]) -> list[dict[str, object]]:
    ui_by_req = by_requirement(selection, "ui_design")
    tc_by_req = by_requirement(selection, "tc_template")
    ts_by_req = by_requirement(selection, "ts_template")
    requirement_ids = sorted(set(ui_by_req) & set(tc_by_req) & set(ts_by_req))
    return [
        {
            "requirement_id": requirement_id,
            "ui_design_path": ui_by_req[requirement_id]["path"],
            "tc_template_path": tc_by_req[requirement_id]["path"],
            "ts_template_path": ts_by_req[requirement_id]["path"],
        }
        for requirement_id in requirement_ids
    ]


def by_requirement(selection: dict[str, dict[str, object]], key: str) -> dict[str, dict[str, object]]:
    value = selection.get(key, {}).get("by_requirement", {})
    return value if isinstance(value, dict) else {}


def build_missing_requirement_report(selection: dict[str, dict[str, object]]) -> list[dict[str, object]]:
    role_labels = {
        "ui_design": "사용자인터페이스설계서",
        "tc_template": "단위시험케이스",
        "ts_template": "통합시험시나리오",
    }
    indexes = {
        key: by_requirement(selection, key)
        for key in role_labels
    }
    requirement_ids = sorted(set().union(*(set(index) for index in indexes.values())))
    return [
        {
            "requirement_id": requirement_id,
            "missing": [
                label
                for key, label in role_labels.items()
                if requirement_id not in indexes[key]
            ],
        }
        for requirement_id in requirement_ids
        if any(requirement_id not in index for index in indexes.values())
    ]


def build_role_counts(selection: dict[str, dict[str, object]]) -> dict[str, int]:
    return {
        key: len(by_requirement(selection, key))
        for key in ("ui_design", "tc_template", "ts_template")
    }


def build_matching_failure_payload(
        *,
        request_id: str,
        dump_root: Path,
        selection: dict[str, dict[str, object]],
        message: str,
) -> dict[str, object]:
    return {
        "ok": False,
        "request_id": request_id,
        "dump_root": str(dump_root),
        "error": message,
        "tc_count": 0,
        "ts_count": 0,
        "requirement_count": 0,
        "processed_requirement_count": 0,
        "failed_requirement_count": 0,
        "role_counts": build_role_counts(selection),
        "source_files": serialize_selection(selection),
        "requirement_items": [],
        "missing_requirements": build_missing_requirement_report(selection),
        "placed_files": [],
        "files": [],
    }


def select_best_file(
        root: Path,
        files: list[Path],
        *,
        suffixes: set[str],
        keywords: tuple[str, ...],
        exclude_keywords: tuple[str, ...] = (),
) -> dict[str, object]:
    scored: list[tuple[int, int, Path]] = []
    for path in files:
        if path.suffix.lower() not in suffixes:
            continue

        text = searchable_text(root, path)
        score = score_keywords(text, keywords) - score_keywords(text, exclude_keywords)
        if score <= 0:
            continue

        scored.append((score, -len(str(path)), path))

    scored.sort(key=lambda item: (item[0], item[1], str(item[2]).casefold()), reverse=True)
    candidates = [
        {
            "path": str(path),
            "score": score,
        }
        for score, _length_score, path in scored[:5]
    ]
    return {
        "path": str(scored[0][2]) if scored else "",
        "score": scored[0][0] if scored else 0,
        "candidates": candidates,
    }


def searchable_text(root: Path, path: Path) -> str:
    try:
        relative = path.relative_to(root)
    except ValueError:
        relative = path
    return normalize_search_text(f"{relative.as_posix()} {path.stem}")


def normalize_search_text(value: str) -> str:
    return re.sub(r"[\s_\-()\[\]{}./\\]+", "", value).casefold()


def score_keywords(text: str, keywords: tuple[str, ...]) -> int:
    score = 0
    for keyword in keywords:
        normalized = normalize_search_text(keyword)
        if normalized and normalized in text:
            score += 100 + len(normalized)
    return score


def require_selected(selection: dict[str, dict[str, object]], key: str, label: str) -> Path:
    selected = Path(str(selection.get(key, {}).get("path") or ""))
    if not selected.exists() or not selected.is_file():
        raise ValueError(f"{label}를 결과 폴더에서 찾지 못했습니다.")
    return selected


def find_generated_file(payload: dict[str, object], *, suffix: str, kind: str) -> Path | None:
    files = payload.get("files")
    if not isinstance(files, list):
        return None

    normalized_suffix = suffix.lower()
    for item in files:
        if not isinstance(item, dict):
            continue
        path = Path(str(item.get("path") or ""))
        item_kind = str(item.get("kind") or "").lower()
        if path.exists() and path.suffix.lower() == normalized_suffix and item_kind == kind:
            return path

    for item in files:
        if not isinstance(item, dict):
            continue
        path = Path(str(item.get("path") or ""))
        if path.exists() and path.suffix.lower() == normalized_suffix:
            return path

    return None


def place_generated_file(
        source_path: Path,
        target_path: Path,
        kind: str,
        label: str,
        requirement_id: str = "",
) -> dict[str, object]:
    # 기존 산출물은 같은 폴더의 bak 하위로 이동하고 생성 파일을 기존 위치/이름에 맞춰 넣는다.
    source_path = Path(source_path)
    target_path = Path(target_path)
    if not source_path.exists() or not source_path.is_file():
        raise FileNotFoundError(f"생성 파일을 찾지 못했습니다: {source_path}")

    backup_path = ""
    if target_path.exists():
        backup_dir = target_path.parent / "bak"
        backup_dir.mkdir(exist_ok=True)
        backup_target = unique_file_path(backup_dir / target_path.name)
        shutil.move(str(target_path), str(backup_target))
        backup_path = str(backup_target)

    target_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(source_path), str(target_path))
    return {
        "kind": kind,
        "label": label,
        "requirement_id": requirement_id,
        "path": str(target_path),
        "backup_path": backup_path,
    }


def unique_file_path(path: Path) -> Path:
    if not path.exists():
        return path

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = path.with_name(f"{path.stem}_{timestamp}{path.suffix}")
    index = 1
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}_{timestamp}_{index}{path.suffix}")
        index += 1
    return candidate


def safe_requirement_dirname(requirement_id: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", requirement_id) or "requirement"


def serialize_selection(selection: dict[str, dict[str, object]]) -> list[dict[str, object]]:
    labels = {
        "ui_design": "사용자인터페이스설계서",
        "tc_template": "단위시험케이스",
        "ts_template": "통합시험시나리오",
    }
    items: list[dict[str, object]] = []
    for key, value in selection.items():
        by_req = value.get("by_requirement", {})
        if not isinstance(by_req, dict):
            continue
        for requirement_id, selected in sorted(by_req.items()):
            if not isinstance(selected, dict):
                continue
            items.append({
                "role": key,
                "label": labels.get(key, key),
                "requirement_id": requirement_id,
                "path": str(selected.get("path") or ""),
                "score": selected.get("score", 0),
                "candidates": selected.get("candidates", []),
            })
    return items
