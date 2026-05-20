# 산출물 매핑 확인 웹 화면을 제공하고, 실제 처리는 기능별 모듈로 넘깁니다.
from __future__ import annotations

import argparse
import json
import mimetypes
from pathlib import Path
import re
import shutil
import sys
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, unquote, urlparse
from uuid import uuid4

from app_runtime import (
    RESULT_DIR,
    WEB_DIR,
    TEMP_DIR,
    TS_TEMPLATE_PATH,
    cleanup_runtime,
    ensure_runtime_dirs,
    log_event,
    parse_json_object,
    read_runtime_env,
    remove_runtime_path,
    resolve_runtime_model,
    resolve_runtime_ollama_chat_url,
    runtime_mode_payload,
)
from output_file_check.folder_workflow import run_web_check, run_web_folder_apply
from web_uploads import parse_multipart_items, safe_upload_filename
from document_update.metadata_workflow import (
    apply_metadata_to_existing_dump,
    run_metadata_apply,
    run_metadata_preview,
    resolve_existing_dump_root,
    save_metadata_inputs,
    save_metadata_required_files,
    split_excluded_paths,
)
from qa_generation.generate_tc import (
    call_ollama,
    extract_process_flow_steps,
    extract_screen_blocks,
    extract_text_from_pdf,
    generate_test_cases,
)
from qa_generation.generate_ts import generate_test_scenarios
from qa_generation.folder_pipeline import QaFolderMatchingError, run_folder_qa_pipeline
from qa_generation.generate_ts import extract_req_mapping_from_pdf, extract_unit_test_from_excel


RESULT_FILES: dict[str, Path] = {}
RESULT_DOWNLOAD_NAMES: dict[str, str] = {}
RESULT_DELETE_AFTER_DOWNLOAD: dict[str, bool] = {}
RESULT_CLEANUP_ROOTS: dict[str, Path] = {}
TS_SET_KEY_PATTERN = re.compile(r"\bSFR-[A-Z0-9]+-\d{3}\b", re.IGNORECASE)


def attach_folder_download(payload: dict[str, object]) -> None:
    # 폴더 반영 결과를 브라우저에서 받을 수 있도록 덤프 폴더를 ZIP으로 묶어 다운로드 토큰을 붙인다.
    dump_root = Path(str(payload.get("dump_root") or ""))
    if not dump_root.is_dir():
        log_event("folder_download.missing_dump", dump_root=str(dump_root))
        return

    token = uuid4().hex
    download_name = f"{dump_root.name}_결과.zip"
    zip_path = Path(shutil.make_archive(str(RESULT_DIR / f"{token}_{dump_root.name}_결과"), "zip", dump_root.parent, dump_root.name))

    RESULT_FILES[token] = zip_path
    RESULT_DOWNLOAD_NAMES[token] = download_name
    payload["download_url"] = f"/download/{token}"
    payload["download_name"] = download_name
    log_event(
        "folder_download.ready",
        token=token,
        dump_root=str(dump_root),
        zip_path=str(zip_path),
    )


def attach_file_downloads(
        payload: dict[str, object],
        *,
        delete_after_download: bool = False,
        cleanup_root: Path | None = None,
) -> None:
    # 생성 모듈이 반환한 파일 목록에 브라우저 다운로드 URL을 붙인다.
    """payload["files"]의 실제 파일 경로를 다운로드 토큰으로 등록한다.

    delete_after_download=True이면 다운로드 응답을 보낸 뒤 해당 파일을 삭제한다.
    cleanup_root가 전달되면 같은 요청의 출력 파일을 모두 받은 뒤 요청 임시 폴더까지 정리한다.
    """
    files = payload.get("files")
    if not isinstance(files, list):
        return

    for item in files:
        if not isinstance(item, dict):
            continue

        path = Path(str(item.get("path") or ""))
        if not path.exists() or not path.is_file():
            continue

        token = uuid4().hex
        download_name = str(item.get("name") or path.name)

        RESULT_FILES[token] = path
        RESULT_DOWNLOAD_NAMES[token] = download_name
        RESULT_DELETE_AFTER_DOWNLOAD[token] = delete_after_download
        if cleanup_root is not None:
            RESULT_CLEANUP_ROOTS[token] = cleanup_root
        item["download_url"] = f"/download/{token}"
        item["download_name"] = download_name

    payload["download_files"] = [
        item
        for item in files
        if isinstance(item, dict) and item.get("download_url")
    ]


def cleanup_sent_download(token: str, path: Path) -> None:
    # 임시 QA 결과 파일은 다운로드가 끝난 뒤 토큰과 파일을 함께 정리한다.
    if not RESULT_DELETE_AFTER_DOWNLOAD.pop(token, False):
        return

    RESULT_FILES.pop(token, None)
    RESULT_DOWNLOAD_NAMES.pop(token, None)
    cleanup_root = RESULT_CLEANUP_ROOTS.pop(token, None)
    remove_runtime_path(path)

    if cleanup_root is None:
        return

    output_dir = path.parent
    try:
        if output_dir.exists() and any(output_dir.iterdir()):
            return
    except OSError:
        return

    remove_runtime_path(cleanup_root)
    log_event("download.cleaned", token=token, cleanup_root=str(cleanup_root))


def save_uploaded_file(
        temp_dir: Path,
        file_items: dict[str, list[tuple[str, bytes]]],
        field_name: str,
        fallback_name: str,
        allowed_suffixes: set[str],
        *,
        required: bool = True,
) -> tuple[Path, str] | None:
    # multipart 업로드에서 지정 필드의 첫 번째 파일을 임시 폴더에 저장한다.
    """필수 여부와 확장자를 검증한 뒤 안전한 파일명으로 저장한다.

    저장 경로와 브라우저가 보낸 원본 파일명을 함께 반환한다.
    파일이 선택 사항이고 업로드되지 않았으면 None을 반환한다.
    """
    items = file_items.get(field_name) or []
    if not items:
        if required:
            raise ValueError(f"{field_name} 파일을 업로드해주세요.")
        return None
    
    filename, payload = items[0]
    if not payload:
        if required:
            raise ValueError(f"{field_name} 파일이 비어있습니다.")
        return None
    
    suffix = Path(filename).suffix.lower()
    if suffix not in allowed_suffixes:
        allowed = ", ".join(allowed_suffixes)
        raise ValueError(f"{field_name} 파일 형식은 {allowed}만 허용됩니다.")
    
    safe_name = safe_upload_filename(filename, field_name, Path(fallback_name).suffix)
    path = temp_dir / safe_name
    path.write_bytes(payload)
    return path, filename


def uploaded_stem(filename: str, field_name: str) -> str:
    # 다운로드 이름에 쓸 원본 파일명 stem을 안전한 형태로 만든다.
    suffix = Path(filename).suffix
    safe_name = safe_upload_filename(filename, field_name, suffix)
    return Path(safe_name).stem or field_name


def apply_download_stem(payload: dict[str, object], stem: str) -> None:
    # 실제 저장 파일명은 유지하고 브라우저 다운로드 파일명만 업로드 원본 stem으로 맞춘다.
    files = payload.get("files")
    if not isinstance(files, list):
        return

    for item in files:
        if not isinstance(item, dict):
            continue

        path = Path(str(item.get("path") or ""))
        suffix = path.suffix or f".{item.get('kind') or 'file'}"
        item["name"] = f"{stem}{suffix}"


def unique_download_name(existing_names: set[str], name: str) -> str:
    path = Path(name)
    candidate = name
    index = 2
    while candidate.casefold() in existing_names:
        candidate = f"{path.stem}_{index}{path.suffix}"
        index += 1
    existing_names.add(candidate.casefold())
    return candidate


def extract_ts_set_key(text: str) -> str:
    match = TS_SET_KEY_PATTERN.search(text or "")
    return match.group(0).upper() if match else ""


def save_uploaded_items(
        temp_dir: Path,
        file_items: dict[str, list[tuple[str, bytes]]],
        field_name: str,
        fallback_name: str,
        allowed_suffixes: set[str],
) -> list[dict[str, object]]:
    saved: list[dict[str, object]] = []
    items = file_items.get(field_name) or []
    for index, (filename, payload) in enumerate(items, start=1):
        if not payload:
            continue
        suffix = Path(filename).suffix.lower()
        if suffix not in allowed_suffixes:
            continue
        safe_name = safe_upload_filename(filename, f"{field_name}_{index}", Path(fallback_name).suffix)
        path = temp_dir / f"{field_name}_{index}_{safe_name}"
        path.write_bytes(payload)
        saved.append(
            {
                "field": field_name,
                "name": filename or safe_name,
                "safe_name": safe_name,
                "stem": Path(safe_name).stem,
                "path": path,
                "key": extract_ts_set_key(safe_name),
            }
        )
    if not saved:
        raise ValueError(f"{field_name} 파일을 선택하세요.")
    return saved


def analyze_ts_tc_file(item: dict[str, object]) -> dict[str, object]:
    path = Path(str(item["path"]))
    unit_test_data = extract_unit_test_from_excel(path)
    screen_ids = sorted(
        {
            str(row.get("화면_ID", "")).strip()
            for row in unit_test_data
            if str(row.get("화면_ID", "")).strip()
        }
    )
    return {
        **item,
        "unit_test_data": unit_test_data,
        "screen_ids": screen_ids,
        "row_count": len(unit_test_data),
        "key": str(item.get("key") or "") or extract_ts_set_key(" ".join(screen_ids)),
    }


def analyze_ts_ui_file(item: dict[str, object]) -> dict[str, object]:
    path = Path(str(item["path"]))
    req_mapping = extract_req_mapping_from_pdf(path)
    screen_ids = sorted(req_mapping.keys())
    req_ids = sorted(set(req_mapping.values()))
    return {
        **item,
        "req_mapping": req_mapping,
        "screen_ids": screen_ids,
        "req_count": len(req_mapping),
        "key": str(item.get("key") or "") or extract_ts_set_key(" ".join(req_ids)),
    }


def rule_match_ts_sets(
        tc_items: list[dict[str, object]],
        ui_items: list[dict[str, object]],
) -> tuple[list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    matches: list[dict[str, object]] = []
    used_ui: set[str] = set()

    for tc_item in tc_items:
        tc_screens = set(tc_item.get("screen_ids") or [])
        best_ui = None
        best_score = 0
        best_reason = ""
        for ui_item in ui_items:
            ui_name = str(ui_item.get("name") or "")
            if ui_name in used_ui:
                continue
            ui_screens = set(ui_item.get("screen_ids") or [])
            score = 0
            reason = ""
            if tc_item.get("key") and tc_item.get("key") == ui_item.get("key"):
                score += 1000
                reason = f"파일명 키 일치: {tc_item['key']}"
            overlap = len(tc_screens & ui_screens)
            if overlap:
                score += overlap * 10
                reason = f"{reason}, 화면ID {overlap}개 일치" if reason else f"화면ID {overlap}개 일치"
            if score > best_score:
                best_score = score
                best_ui = ui_item
                best_reason = reason

        if best_ui is None:
            continue

        used_ui.add(str(best_ui.get("name") or ""))
        matches.append(
            {
                "tc": tc_item,
                "ui": best_ui,
                "confidence": min(1.0, best_score / 1000),
                "reason": best_reason,
            }
        )

    matched_tc_names = {str(match["tc"].get("name") or "") for match in matches}
    unmatched_tcs = [item for item in tc_items if str(item.get("name") or "") not in matched_tc_names]
    unmatched_uis = [item for item in ui_items if str(item.get("name") or "") not in used_ui]
    return matches, unmatched_tcs, unmatched_uis


def ai_refine_ts_sets(
        rule_matches: list[dict[str, object]],
        unmatched_tcs: list[dict[str, object]],
        unmatched_uis: list[dict[str, object]],
        model_name: str,
        ollama_url: str,
) -> dict[str, object]:
    if not ollama_url or not model_name:
        return {"summary": "AI 매칭을 건너뛰고 규칙 기반 매칭을 사용했습니다.", "sets": []}

    def brief_file(item: dict[str, object]) -> dict[str, object]:
        return {
            "name": item.get("name"),
            "key": item.get("key"),
            "screen_count": len(item.get("screen_ids") or []),
            "sample_screen_ids": list(item.get("screen_ids") or [])[:8],
        }

    prompt = {
        "task": "통합시험 시나리오 생성을 위한 단위시험 케이스 XLSX와 사용자인터페이스설계서 PDF 세트를 매칭하세요.",
        "rules": [
            "같은 SFR 키가 있으면 같은 세트일 가능성이 높습니다.",
            "SFR 키가 없거나 애매하면 화면ID 교집합을 기준으로 판단합니다.",
            "확신이 낮은 파일은 억지로 매칭하지 말고 unmatched로 둡니다.",
        ],
        "rule_matches": [
            {
                "tc_file": match["tc"].get("name"),
                "ui_file": match["ui"].get("name"),
                "reason": match.get("reason", ""),
            }
            for match in rule_matches
        ],
        "unmatched_tcs": [brief_file(item) for item in unmatched_tcs],
        "unmatched_uis": [brief_file(item) for item in unmatched_uis],
        "output_format": {
            "summary": "한 문장 요약",
            "sets": [
                {
                    "tc_file": "TC 파일명",
                    "ui_file": "UI PDF 파일명",
                    "confidence": 0.0,
                    "reason": "매칭 근거",
                }
            ],
            "risks": ["확인 필요 사항"],
        },
    }

    try:
        raw = call_ollama(
            ollama_url,
            model_name,
            "당신은 QA 산출물 입력 파일을 세트로 매칭하는 분석가입니다. JSON 객체만 출력하세요.",
            json.dumps(prompt, ensure_ascii=False),
            num_predict=2048,
            timeout=90,
        )
        parsed = parse_json_object(raw)
        return parsed if parsed else {"summary": "AI 매칭 응답을 해석하지 못했습니다.", "sets": []}
    except Exception as exc:
        return {"summary": "AI 매칭에 실패해 규칙 기반 매칭을 사용했습니다.", "sets": [], "ai_error": str(exc)}


def analyze_tc_pdf(pdf_path: Path, source_name: str, model_name: str, ollama_url: str) -> dict[str, object]:
    analysis: dict[str, object] = {
        "summary": "",
        "quality": "warning",
        "screen_count": 0,
        "screens": [],
        "risks": [],
        "recommendations": [],
    }

    try:
        extracted_text = extract_text_from_pdf(pdf_path)
        screen_blocks = extract_screen_blocks(extracted_text)
    except Exception as exc:
        analysis["summary"] = "PDF 텍스트 추출 또는 화면 분석에 실패했습니다."
        analysis["quality"] = "poor"
        analysis["risks"] = [str(exc)]
        return analysis

    screens: list[dict[str, object]] = []
    for block in screen_blocks:
        block_text = str(block.get("text") or "")
        flow_steps = extract_process_flow_steps(block_text)
        screens.append(
            {
                "screen_id": block.get("screen_id") or "",
                "unit_test_id": block.get("unit_test_id") or "",
                "process_step_count": len(flow_steps),
            }
        )

    risks: list[str] = []
    if not extracted_text.strip():
        risks.append("PDF에서 텍스트를 추출하지 못했습니다.")
    if not screens:
        risks.append("화면ID를 찾지 못했습니다.")
    no_step_screens = [screen["screen_id"] for screen in screens if not screen.get("process_step_count")]
    if no_step_screens:
        risks.append(f"처리흐름이 없는 화면 {len(no_step_screens)}개가 있습니다.")

    analysis.update(
        {
            "summary": f"{source_name}에서 화면 {len(screens)}개를 찾았습니다.",
            "quality": "poor" if not screens else ("warning" if risks else "good"),
            "screen_count": len(screens),
            "screens": screens[:30],
            "risks": risks,
            "recommendations": [],
            "_extracted_text": extracted_text,
            "_screen_blocks": screen_blocks,
        }
    )

    if not screens:
        return analysis

    screen_brief = "\n".join(
        f"- {screen['screen_id']} / {screen['unit_test_id']} / 처리흐름 {screen['process_step_count']}개"
        for screen in screens[:20]
    )
    prompt = f"""
다음 사용자인터페이스설계서 PDF의 사전 분석 결과를 보고 단위시험 케이스 생성 관점으로 요약하세요.
반드시 JSON 객체만 출력하세요.

파일명: {source_name}
추출 화면 수: {len(screens)}
화면 목록:
{screen_brief}

현재 감지된 위험:
{chr(10).join(f"- {risk}" for risk in risks) if risks else "- 없음"}

출력 형식:
{{
  "summary": "한 문장 요약",
  "quality": "good 또는 warning 또는 poor",
  "risks": ["위험요인"],
  "recommendations": ["생성 전 확인 또는 보완 권고"]
}}
"""
    try:
        raw = call_ollama(
            ollama_url,
            model_name,
            "당신은 QA 산출물 생성 전 입력 문서를 점검하는 분석가입니다.",
            prompt,
            num_predict=1024,
            timeout=60,
        )
        ai = parse_json_object(raw)
        if isinstance(ai.get("summary"), str) and ai["summary"].strip():
            analysis["summary"] = ai["summary"].strip()
        if ai.get("quality") in {"good", "warning", "poor"}:
            analysis["quality"] = ai["quality"]
        if isinstance(ai.get("risks"), list):
            analysis["risks"] = [str(item) for item in ai["risks"] if str(item).strip()]
        if isinstance(ai.get("recommendations"), list):
            analysis["recommendations"] = [
                str(item) for item in ai["recommendations"] if str(item).strip()
            ]
    except Exception as exc:
        analysis["ai_error"] = str(exc)

    return analysis


def parse_json_fields(content_type: str, body: bytes) -> dict[str, str]:
    # JSON 요청을 기존 runtime_ai_settings가 쓰는 문자열 필드 dict로 바꾼다.
    if "application/json" not in content_type.lower():
        fields, _files = parse_multipart_items(content_type, body)
        return fields

    raw_payload = json.loads(body.decode("utf-8") or "{}")
    if not isinstance(raw_payload, dict):
        raise ValueError("JSON 요청 본문은 객체여야 합니다.")

    return {
        str(key): "" if value is None else str(value)
        for key, value in raw_payload.items()
    }


def runtime_ai_settings(fields: dict[str, str]) -> tuple[str, str]:
    # 요청 필드와 .env를 합쳐 QA 생성에 사용할 모델명과 Ollama chat URL을 결정한다.
    """프론트에서 넘어온 값이 있으면 우선 사용하고, 없으면 .env 설정을 사용한다."""
    env = read_runtime_env()
    model_name = fields.get("model_name") or resolve_runtime_model(env)
    ollama_url = fields.get("ollama_url") or resolve_runtime_ollama_chat_url(env)

    if not ollama_url:
        raise ValueError("OLLAMA_BASE_URL이 설정되어 있지 않습니다.")
    
    return model_name, ollama_url


class WebHandler(BaseHTTPRequestHandler):
    server_version = "OutputMappingHTTP/1.0"

    def do_GET(self) -> None:
        # 정적 화면, 런타임 모드 API, 다운로드 요청을 처리한다.
        request_path = urlparse(self.path).path
        if request_path == "/api/runtime-mode":
            self.send_json(runtime_mode_payload())
            return

        if request_path in {"/", "/check", "/check.html"}:
            self.serve_file(WEB_DIR / "check.html")
            return

        if request_path in {"/qa", "/qa.html"}:
            self.serve_file(WEB_DIR / "qa.html")
            return

        if request_path in {"/metadata", "/metadata.html"}:
            self.serve_file(WEB_DIR / "metadata.html")
            return

        if request_path.startswith("/static/"):
            relative_path = unquote(request_path.removeprefix("/static/"))
            self.serve_file(WEB_DIR / "static" / relative_path)
            return

        if request_path.startswith("/download/"):
            self.serve_download(request_path.removeprefix("/download/"))
            return

        self.send_error(404)

    def do_POST(self) -> None:
        # 폴더 검사, 폴더 덤프 반영 POST 요청을 분기한다.
        request_path = urlparse(self.path).path
        
        if request_path == "/api/folder-apply":
            self.handle_folder_apply_post()
            return

        if request_path == "/api/check":
            self.handle_check_post()
            return
        
        if request_path == "/api/generate-tc":
            self.handle_generate_tc_post()
            return

        if request_path == "/api/generate-ts":
            self.handle_generate_ts_post()
            return

        if request_path == "/api/run-qa-folder":
            self.handle_run_qa_folder_post()
            return

        if request_path == "/api/metadata-preview":
            self.handle_metadata_preview_post()
            return

        if request_path == "/api/metadata-apply":
            self.handle_metadata_apply_post()
            return

        self.send_error(404)

    def handle_metadata_preview_post(self) -> None:
        temp_dir: Path | None = None
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)
            log_event("metadata.preview.post", content_length=content_length, content_type=content_type)
            fields, file_items = parse_multipart_items(content_type, body)
            request_id = uuid4().hex[:8]
            temp_dir = TEMP_DIR / f"metadata-preview-{request_id}"
            temp_dir.mkdir(parents=True, exist_ok=True)
            if fields.get("dump_root", "").strip():
                wbs_path, standard_path = save_metadata_required_files(temp_dir, file_items)
                folder_root = resolve_existing_dump_root(fields.get("dump_root", ""))
            else:
                wbs_path, standard_path, folder_root, _uploaded = save_metadata_inputs(temp_dir, file_items)
            payload = run_metadata_preview(wbs_path, standard_path, folder_root, request_id)
            self.send_json(payload)
        except Exception as exc:
            log_event("metadata.preview.error", error=str(exc), traceback=traceback.format_exc())
            self.send_json({"ok": False, "error": str(exc), "targets": []}, status=400)
        finally:
            if temp_dir is not None:
                remove_runtime_path(temp_dir)

    def handle_metadata_apply_post(self) -> None:
        temp_dir: Path | None = None
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)
            log_event("metadata.apply.post", content_length=content_length, content_type=content_type)
            fields, file_items = parse_multipart_items(content_type, body)
            request_id = uuid4().hex[:8]
            temp_dir = TEMP_DIR / f"metadata-apply-{request_id}"
            temp_dir.mkdir(parents=True, exist_ok=True)
            if fields.get("dump_root", "").strip():
                wbs_path, standard_path = save_metadata_required_files(temp_dir, file_items)
                folder_root = resolve_existing_dump_root(fields.get("dump_root", ""))
                payload = apply_metadata_to_existing_dump(
                    wbs_path,
                    standard_path,
                    folder_root,
                    request_id,
                    split_excluded_paths(fields.get("excluded_paths", "")),
                    temp_parent=temp_dir,
                )
            else:
                wbs_path, standard_path, folder_root, _uploaded = save_metadata_inputs(temp_dir, file_items)
                payload = run_metadata_apply(
                    wbs_path,
                    standard_path,
                    folder_root,
                    request_id,
                    split_excluded_paths(fields.get("excluded_paths", "")),
                )
                attach_folder_download(payload)
            self.send_json(payload, status=200 if payload.get("ok") else 400)
        except Exception as exc:
            log_event("metadata.apply.error", error=str(exc), traceback=traceback.format_exc())
            self.send_json({"ok": False, "error": str(exc), "apply_items": []}, status=400)
        finally:
            if temp_dir is not None:
                remove_runtime_path(temp_dir)

    def handle_check_post(self) -> None:
        # /api/check 요청 본문을 파싱하고 폴더 매칭만 실행한다.
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)
            log_event("http.post", path=self.path, content_length=content_length, content_type=content_type)
            fields, file_items = parse_multipart_items(content_type, body)
            payload = run_web_check(fields, file_items)
            self.send_json(payload)
        except Exception as exc:
            log_event("http.post.error", path=self.path, error=str(exc))
            self.send_json({"error": str(exc)}, status=400)

    def handle_folder_apply_post(self) -> None:
        # /api/folder-apply 요청 본문을 파싱하고 덤프 반영까지 실행한다.
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)
            log_event("http.post", path=self.path, content_length=content_length, content_type=content_type)
            fields, file_items = parse_multipart_items(content_type, body)
            payload = run_web_folder_apply(fields, file_items)
            payload["output_mode"] = "folder"
            if isinstance(payload.get("dump_root"), str):
                log_event("folder_result.ready", dump_root=str(payload["dump_root"]))
            self.send_json(payload)
        except Exception as exc:
            log_event("http.post.error", path=self.path, error=str(exc))
            self.send_json({"error": str(exc)}, status=400)

    def handle_generate_tc_post(self) -> None:
        # 단위시험 케이스 생성 요청을 처리한다.
        """업로드된 HWPX 양식과 UI 설계서 PDF를 임시 폴더에 저장한 뒤 TC 생성 모듈을 실행한다.

        생성 결과는 요청별 output 폴더에 두고 다운로드 토큰을 붙인다.
        업로드 원본은 생성 직후 삭제하고, 결과 파일은 다운로드 후 정리한다.
        """
        temp_dir: Path | None = None
        preserve_temp_dir = False
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)

            fields, file_items = parse_multipart_items(content_type, body)
            model_name, ollama_url = runtime_ai_settings(fields)

            request_id = uuid4().hex[:8]
            temp_dir = TEMP_DIR / f"qa-tc-{request_id}"
            temp_dir.mkdir(parents=True, exist_ok=True)
            output_dir = temp_dir / "output"
            output_dir.mkdir()

            template_hwpx_path, _template_hwpx_name = save_uploaded_file(
                temp_dir,
                file_items,
                "template_hwpx",
                "template.hwpx",
                {".hwpx"},
            )
            ui_pdf_items = file_items.get("ui_pdf") or []
            if not ui_pdf_items:
                raise ValueError("사용자인터페이스 설계서 문서를 선택하세요.")

            log_event(
                "qa.tc.start",
                request_id=request_id,
                model=model_name,
                files={name: len(items) for name, items in file_items.items()},
            )

            all_files: list[dict[str, object]] = []
            source_results: list[dict[str, object]] = []
            download_names: set[str] = set()
            total_count = 0

            for index, (pdf_filename, pdf_payload) in enumerate(ui_pdf_items, start=1):
                source_name = pdf_filename or f"ui_pdf_{index}.pdf"
                pdf_path: Path | None = None
                pdf_output_dir: Path | None = None
                analysis: dict[str, object] | None = None
                try:
                    if not pdf_payload:
                        raise ValueError("빈 문서 파일입니다.")

                    source_suffix = Path(source_name).suffix.lower()
                    if source_suffix not in {".hwp", ".hwpx", ".pdf"}:
                        raise ValueError("HWP, HWPX, PDF 파일만 업로드할 수 있습니다.")

                    safe_pdf_name = safe_upload_filename(source_name, f"ui_pdf_{index}", source_suffix or ".pdf")
                    pdf_stem = Path(safe_pdf_name).stem
                    pdf_path = temp_dir / f"ui_pdf_{index}_{safe_pdf_name}"
                    pdf_path.write_bytes(pdf_payload)
                    analysis = analyze_tc_pdf(pdf_path, source_name, model_name, ollama_url)
                    extracted_text = None
                    screen_blocks = None
                    if isinstance(analysis, dict):
                        extracted_text_value = analysis.pop("_extracted_text", None)
                        if isinstance(extracted_text_value, str):
                            extracted_text = extracted_text_value
                        screen_blocks_value = analysis.pop("_screen_blocks", None)
                        if isinstance(screen_blocks_value, list):
                            screen_blocks = screen_blocks_value

                    pdf_output_dir = output_dir / f"{index:03d}_{pdf_stem}"
                    pdf_output_dir.mkdir(parents=True, exist_ok=True)

                    item_payload = generate_test_cases(
                        pdf_path=pdf_path,
                        model_name=model_name,
                        ollama_url=ollama_url,
                        output_dir=pdf_output_dir,
                        template_path=template_hwpx_path,
                        extracted_text=extracted_text,
                        screen_blocks=screen_blocks,
                    )

                    item_count = int(item_payload.get("count") or 0)
                    total_count += item_count
                    item_files = item_payload.get("files") if isinstance(item_payload.get("files"), list) else []

                    for file_item in item_files:
                        if not isinstance(file_item, dict):
                            continue
                        path = Path(str(file_item.get("path") or ""))
                        suffix = path.suffix or f".{file_item.get('kind') or 'file'}"
                        download_name = unique_download_name(
                            download_names,
                            f"{pdf_stem}{suffix}",
                        )
                        target_path = output_dir / download_name
                        if path.exists() and path.resolve() != target_path.resolve():
                            path.replace(target_path)
                            file_item["path"] = str(target_path)
                        file_item["name"] = download_name
                        file_item["source_pdf"] = source_name
                        all_files.append(file_item)

                    source_results.append(
                        {
                            "source_pdf": source_name,
                            "ok": bool(item_payload.get("ok")),
                            "count": item_count,
                            "file_count": len(item_files),
                            "error": str(item_payload.get("error") or ""),
                            "analysis": analysis,
                        }
                    )
                except Exception as exc:
                    source_results.append(
                        {
                            "source_pdf": source_name,
                            "ok": False,
                            "count": 0,
                            "file_count": 0,
                            "error": str(exc),
                            "analysis": analysis,
                        }
                    )
                finally:
                    if pdf_path is not None:
                        remove_runtime_path(pdf_path)
                    if pdf_output_dir is not None:
                        remove_runtime_path(pdf_output_dir)

            failed_count = sum(1 for item in source_results if not item.get("ok"))
            payload = {
                "ok": bool(all_files),
                "count": total_count,
                "files": all_files,
                "source_results": source_results,
                "source_count": len(source_results),
                "failed_count": failed_count,
            }
            if not all_files:
                errors = [
                    f"{item.get('source_pdf')}: {item.get('error')}"
                    for item in source_results
                    if item.get("error")
                ]
                payload["error"] = "\n".join(errors) or "단위시험 케이스 생성 결과가 없습니다."
            payload["request_id"] = request_id
            attach_file_downloads(payload, delete_after_download=True, cleanup_root=temp_dir)
            preserve_temp_dir = bool(payload.get("download_files"))
            remove_runtime_path(template_hwpx_path)

            log_event(
                "qa.tc.done",
                request_id=request_id,
                ok=payload.get("ok"),
                count=payload.get("count"),
                source_count=payload.get("source_count"),
                failed_count=payload.get("failed_count"),
                file_count=len(payload.get("download_files") or []),
            )
            self.send_json(payload, status=200 if payload.get("ok") else 400)

        except Exception as exc:
            log_event("qa.tc.error", error=str(exc), traceback=traceback.format_exc())
            self.send_json({"ok": False, "error": str(exc), "files": []}, status=400)
        finally:
            if temp_dir is not None and not preserve_temp_dir:
                remove_runtime_path(temp_dir)

    def handle_run_qa_folder_post(self) -> None:
        # check.html의 결과 폴더를 기준으로 TC/TS 생성과 기존 파일 교체를 한 번에 실행한다.
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)
            if "application/json" in content_type.lower():
                fields = parse_json_fields(content_type, body)
                file_items: dict[str, list[tuple[str, bytes]]] = {}
            else:
                fields, file_items = parse_multipart_items(content_type, body)
            dump_root_value = fields.get("dump_root", "").strip()
            if not dump_root_value:
                raise ValueError("check 결과 폴더 경로를 입력하세요.")
            model_name, ollama_url = runtime_ai_settings(fields)
            request_id = uuid4().hex[:8]

            payload = run_folder_qa_pipeline(
                Path(dump_root_value),
                model_name=model_name,
                ollama_url=ollama_url,
                scenario_form_path=TS_TEMPLATE_PATH,
                ui_design_items=file_items.get("ui_design_files") or [],
                ui_design_root=Path(fields.get("ui_design_root", "")) if fields.get("ui_design_root") else None,
                qa_source_items=file_items.get("qa_source_files") or [],
                qa_source_root=Path(fields.get("qa_source_root", "")) if fields.get("qa_source_root") else None,
                tc_source_root=Path(fields.get("tc_source_root", "")) if fields.get("tc_source_root") else None,
                ts_source_root=Path(fields.get("ts_source_root", "")) if fields.get("ts_source_root") else None,
                request_id=request_id,
            )
            self.send_json(payload, status=200 if payload.get("ok") else 400)
        except QaFolderMatchingError as exc:
            log_event("qa.folder.post.matching_error", error=str(exc), payload=exc.payload)
            self.send_json(exc.payload, status=400)
        except Exception as exc:
            log_event("qa.folder.post.error", error=str(exc), traceback=traceback.format_exc())
            self.send_json({"ok": False, "error": str(exc), "files": []}, status=400)

    def handle_generate_ts_post(self) -> None:
        # 통합시험 시나리오 생성 요청을 처리한다.
        """기존 시나리오 XLSX, 단위시험 케이스 XLSX, UI 설계서 PDF를 받아 TS 생성 모듈을 실행한다.

        생성 결과는 요청별 output 폴더에 두고 다운로드 토큰을 붙인다.
        업로드 원본은 생성 직후 삭제하고, 결과 파일은 다운로드 후 정리한다.
        """
        temp_dir: Path | None = None
        preserve_temp_dir = False
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            body = self.rfile.read(content_length)

            fields, file_items = parse_multipart_items(content_type, body)
            try:
                model_name, ollama_url = runtime_ai_settings(fields)
            except Exception:
                model_name, ollama_url = "", ""

            request_id = uuid4().hex[:8]
            temp_dir = TEMP_DIR / f"qa-ts-{request_id}"
            temp_dir.mkdir(parents=True, exist_ok=True)
            output_dir = temp_dir / "output"
            output_dir.mkdir()

            template_xlsx_path, template_xlsx_name = save_uploaded_file(
                temp_dir,
                file_items,
                "template_xlsx",
                "template.xlsx",
                {".xlsx"},
            )
            tc_xlsx_path, tc_xlsx_name = save_uploaded_file(
                temp_dir,
                file_items,
                "tc_xlsx",
                "test_cases.xlsx",
                {".xlsx"},
            )
            ui_pdf_path, ui_pdf_name = save_uploaded_file(
                temp_dir,
                file_items,
                "ui_pdf",
                "ui.pdf",
                {".hwp", ".hwpx", ".pdf"},
            )
            tc_items = save_uploaded_items(temp_dir, file_items, "tc_xlsx", "test_cases.xlsx", {".xlsx"})
            ui_items = save_uploaded_items(temp_dir, file_items, "ui_pdf", "ui.pdf", {".hwp", ".hwpx", ".pdf"})

            log_event(
                "qa.ts.start",
                request_id=request_id,
                files={name: len(items) for name, items in file_items.items()},
            )

            analyzed_tcs: list[dict[str, object]] = []
            analyzed_uis: list[dict[str, object]] = []
            pre_errors: list[dict[str, object]] = []

            for item in tc_items:
                try:
                    analyzed_tcs.append(analyze_ts_tc_file(item))
                except Exception as exc:
                    pre_errors.append({"file": item.get("name"), "kind": "tc", "error": str(exc)})

            for item in ui_items:
                try:
                    analyzed_uis.append(analyze_ts_ui_file(item))
                except Exception as exc:
                    pre_errors.append({"file": item.get("name"), "kind": "ui", "error": str(exc)})

            rule_matches, unmatched_tcs, unmatched_uis = rule_match_ts_sets(analyzed_tcs, analyzed_uis)
            ai_matching = ai_refine_ts_sets(rule_matches, unmatched_tcs, unmatched_uis, model_name, ollama_url)

            matches = list(rule_matches)
            matched_tc_names = {str(match["tc"].get("name") or "") for match in matches}
            matched_ui_names = {str(match["ui"].get("name") or "") for match in matches}
            tc_by_name = {str(item.get("name") or ""): item for item in analyzed_tcs}
            ui_by_name = {str(item.get("name") or ""): item for item in analyzed_uis}
            for item in ai_matching.get("sets", []) if isinstance(ai_matching.get("sets"), list) else []:
                if not isinstance(item, dict):
                    continue
                tc_name = str(item.get("tc_file") or "")
                ui_name = str(item.get("ui_file") or "")
                if not tc_name or not ui_name or tc_name in matched_tc_names or ui_name in matched_ui_names:
                    continue
                tc_item = tc_by_name.get(tc_name)
                ui_item = ui_by_name.get(ui_name)
                if not tc_item or not ui_item:
                    continue
                matches.append(
                    {
                        "tc": tc_item,
                        "ui": ui_item,
                        "confidence": item.get("confidence", 0.5),
                        "reason": item.get("reason", "AI 매칭"),
                    }
                )
                matched_tc_names.add(tc_name)
                matched_ui_names.add(ui_name)

            download_names: set[str] = set()
            all_files: list[dict[str, object]] = []
            source_results: list[dict[str, object]] = []
            total_count = 0

            for index, match in enumerate(matches, start=1):
                tc_item = match["tc"]
                ui_item = match["ui"]
                set_stem = str(tc_item.get("stem") or Path(str(tc_item.get("name") or f"set_{index}")).stem)
                set_output_dir = output_dir / f"{index:03d}_{set_stem}"
                set_output_dir.mkdir(parents=True, exist_ok=True)
                try:
                    item_payload = generate_test_scenarios(
                        template_xlsx_path=template_xlsx_path,
                        tc_xlsx_path=Path(str(tc_item["path"])),
                        ui_pdf_path=Path(str(ui_item["path"])),
                        output_dir=set_output_dir,
                        form_path=TS_TEMPLATE_PATH,
                        req_mapping=ui_item.get("req_mapping") if isinstance(ui_item.get("req_mapping"), dict) else None,
                        unit_test_data=tc_item.get("unit_test_data") if isinstance(tc_item.get("unit_test_data"), list) else None,
                        log_progress=index == 1,
                    )
                    item_count = int(item_payload.get("count") or 0)
                    total_count += item_count
                    item_files = item_payload.get("files") if isinstance(item_payload.get("files"), list) else []
                    for file_item in item_files:
                        if not isinstance(file_item, dict):
                            continue
                        path = Path(str(file_item.get("path") or ""))
                        suffix = path.suffix or ".xlsx"
                        download_name = unique_download_name(download_names, f"ts_{set_stem}{suffix}")
                        target_path = output_dir / download_name
                        if path.exists() and path.resolve() != target_path.resolve():
                            path.replace(target_path)
                            file_item["path"] = str(target_path)
                        file_item["name"] = download_name
                        file_item["source_tc"] = tc_item.get("name")
                        file_item["source_ui"] = ui_item.get("name")
                        all_files.append(file_item)

                    source_results.append(
                        {
                            "source_pdf": f"{tc_item.get('name')} + {ui_item.get('name')}",
                            "source_tc": tc_item.get("name"),
                            "source_ui": ui_item.get("name"),
                            "ok": bool(item_payload.get("ok")),
                            "count": item_count,
                            "file_count": len(item_files),
                            "error": str(item_payload.get("error") or ""),
                            "analysis": {
                                "summary": match.get("reason") or ai_matching.get("summary") or "세트 매칭 완료",
                                "quality": "good" if item_payload.get("ok") else "warning",
                                "screen_count": len(tc_item.get("screen_ids") or []),
                                "screens": [{"screen_id": screen_id} for screen_id in (tc_item.get("screen_ids") or [])],
                                "risks": item_payload.get("missing_screen_ids", []) if item_payload.get("missing_screen_ids") else [],
                                "recommendations": [],
                            },
                        }
                    )
                except Exception as exc:
                    source_results.append(
                        {
                            "source_pdf": f"{tc_item.get('name')} + {ui_item.get('name')}",
                            "source_tc": tc_item.get("name"),
                            "source_ui": ui_item.get("name"),
                            "ok": False,
                            "count": 0,
                            "file_count": 0,
                            "error": str(exc),
                            "analysis": {
                                "summary": match.get("reason") or "세트 처리 중 오류가 발생했습니다.",
                                "quality": "poor",
                                "screen_count": len(tc_item.get("screen_ids") or []),
                                "screens": [{"screen_id": screen_id} for screen_id in (tc_item.get("screen_ids") or [])],
                                "risks": [str(exc)],
                                "recommendations": [],
                            },
                        }
                    )
                finally:
                    remove_runtime_path(set_output_dir)

            unmatched_tc_names = sorted(set(tc_by_name) - matched_tc_names)
            unmatched_ui_names = sorted(set(ui_by_name) - matched_ui_names)
            for name in unmatched_tc_names:
                source_results.append(
                    {
                        "source_pdf": name,
                        "ok": False,
                        "count": 0,
                        "file_count": 0,
                        "error": "매칭되는 사용자인터페이스설계서 PDF를 찾지 못했습니다.",
                        "analysis": {"summary": "세트 매칭 실패", "quality": "poor", "screen_count": 0, "risks": ["UI PDF 누락"], "recommendations": []},
                    }
                )
            for name in unmatched_ui_names:
                source_results.append(
                    {
                        "source_pdf": name,
                        "ok": False,
                        "count": 0,
                        "file_count": 0,
                        "error": "매칭되는 단위시험 케이스 XLSX를 찾지 못했습니다.",
                        "analysis": {"summary": "세트 매칭 실패", "quality": "poor", "screen_count": 0, "risks": ["TC XLSX 누락"], "recommendations": []},
                    }
                )
            for item in pre_errors:
                source_results.append(
                    {
                        "source_pdf": item.get("file"),
                        "ok": False,
                        "count": 0,
                        "file_count": 0,
                        "error": item.get("error"),
                        "analysis": {"summary": "사전 분석 실패", "quality": "poor", "screen_count": 0, "risks": [item.get("error")], "recommendations": []},
                    }
                )

            ai_risks = ai_matching.get("risks") if isinstance(ai_matching.get("risks"), list) else []
            if ai_matching.get("summary") or ai_risks or ai_matching.get("ai_error"):
                matched_set_names = [
                    f"{match['tc'].get('name')} + {match['ui'].get('name')}"
                    for match in matches
                ]
                source_results.insert(
                    0,
                    {
                        "source_pdf": "AI 세트 매칭 요약",
                        "is_summary": True,
                        "ok": not bool(ai_matching.get("ai_error")),
                        "count": 0,
                        "file_count": 0,
                        "error": str(ai_matching.get("ai_error") or ""),
                        "analysis": {
                            "summary": str(ai_matching.get("summary") or "세트 매칭 요약을 생성했습니다."),
                            "quality": "warning" if ai_matching.get("ai_error") or ai_risks else "good",
                            "metric_label": "세트",
                            "metric_count": len(matches),
                            "screen_count": 0,
                            "risks": [str(item) for item in ai_risks],
                            "recommendations": matched_set_names[:5],
                        },
                    }
                )

            failed_count = sum(1 for item in source_results if not item.get("ok"))
            payload = {
                "ok": bool(all_files),
                "count": total_count,
                "files": all_files,
                "source_results": source_results,
                "set_count": len(matches),
                "failed_count": failed_count,
                "ai_matching": ai_matching,
            }
            if not all_files:
                payload["error"] = "생성 가능한 통합시험 시나리오 세트를 찾지 못했습니다."
            payload["request_id"] = request_id
            attach_file_downloads(payload, delete_after_download=True, cleanup_root=temp_dir)
            preserve_temp_dir = bool(payload.get("download_files"))
            remove_runtime_path(template_xlsx_path)
            for item in [*tc_items, *ui_items]:
                remove_runtime_path(Path(str(item.get("path") or "")))

            log_event(
                "qa.ts.done",
                request_id=request_id,
                ok=payload.get("ok"),
                count=payload.get("count"),
                set_count=payload.get("set_count"),
                failed_count=payload.get("failed_count"),
                file_count=len(payload.get("download_files") or []),
            )
            self.send_json(payload, status=200 if payload.get("ok") else 400)

        except Exception as exc:
            log_event("qa.ts.error", error=str(exc), traceback=traceback.format_exc())
            self.send_json({"ok": False, "error": str(exc), "files": []}, status=400)
        finally:
            if temp_dir is not None and not preserve_temp_dir:
                remove_runtime_path(temp_dir)


    def serve_file(self, path: Path) -> None:
        # web 폴더의 HTML/CSS/JS 정적 파일을 응답한다.
        """임의의 로컬 파일을 읽지 못하도록 웹 정적 파일 폴더 아래 파일만 제공한다."""
        resolved = path.resolve()
        if WEB_DIR.resolve() not in resolved.parents and resolved != WEB_DIR.resolve():
            self.send_error(403)
            return

        if not resolved.exists() or not resolved.is_file():
            self.send_error(404)
            return

        content_type = mimetypes.guess_type(resolved.name)[0] or "application/octet-stream"
        data = resolved.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def serve_download(self, token: str) -> None:
        # 처리 완료된 결과 파일을 토큰으로 찾아 다운로드 응답한다.
        """생성된 결과 문서를 다운로드로 전송한다."""
        path = RESULT_FILES.get(token)
        if not path or not path.exists():
            self.send_error(404)
            return

        data = path.read_bytes()
        download_name = RESULT_DOWNLOAD_NAMES.get(token, path.name)
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(download_name)}")
        self.end_headers()
        self.wfile.write(data)
        log_event("download.sent", token=token, path=str(path), bytes=len(data))
        cleanup_sent_download(token, path)

    def send_json(self, payload: dict[str, object], status: int = 200) -> None:
        # dict payload를 UTF-8 JSON HTTP 응답으로 보낸다.
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format: str, *args: object) -> None:
        # 기본 콘솔 로그 대신 web_app.log 파일에 접근 로그를 남긴다.
        log_event("http.access", client=self.client_address[0], message=format % args)


def parse_args() -> argparse.Namespace:
    # 웹 서버 host/port 명령줄 옵션을 읽는다.
    parser = argparse.ArgumentParser(description="산출물 매핑 확인 웹 도구")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8000)
    return parser.parse_args()


def main() -> None:
    # 작업 폴더를 정리하고 HTTP 서버를 시작한다.
    args = parse_args()
    read_runtime_env()
    ensure_runtime_dirs()
    cleanup_runtime()

    server = ThreadingHTTPServer((args.host, args.port), WebHandler)
    log_event("server.start", url=f"http://{args.host}:{args.port}")
    print(f"http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
