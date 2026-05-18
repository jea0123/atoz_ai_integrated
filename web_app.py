# 산출물 매핑 확인 웹 화면을 제공하고, 실제 처리는 기능별 모듈로 넘깁니다.
from __future__ import annotations

import argparse
import json
import mimetypes
from pathlib import Path
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
from qa_generation.generate_tc import (
    call_ollama,
    extract_process_flow_steps,
    extract_screen_blocks,
    extract_text_from_pdf,
    generate_test_cases,
)
from qa_generation.generate_ts import generate_test_scenarios


RESULT_FILES: dict[str, Path] = {}
RESULT_DOWNLOAD_NAMES: dict[str, str] = {}
RESULT_DELETE_AFTER_DOWNLOAD: dict[str, bool] = {}
RESULT_CLEANUP_ROOTS: dict[str, Path] = {}


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

        self.send_error(404)

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
            attach_folder_download(payload)
            if not fields.get("dump_path") and isinstance(payload.get("dump_root"), str):
                remove_runtime_path(Path(payload["dump_root"]))
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
                raise ValueError("사용자인터페이스 설계서 PDF를 선택하세요.")

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
                        raise ValueError("빈 PDF 파일입니다.")

                    if Path(source_name).suffix.lower() != ".pdf":
                        raise ValueError("PDF 파일만 업로드할 수 있습니다.")

                    safe_pdf_name = safe_upload_filename(source_name, f"ui_pdf_{index}", ".pdf")
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

            _, file_items = parse_multipart_items(content_type, body)

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
                {".pdf"},
            )

            log_event(
                "qa.ts.start",
                request_id=request_id,
                files={name: len(items) for name, items in file_items.items()},
            )

            payload = generate_test_scenarios(
                template_xlsx_path=template_xlsx_path,
                tc_xlsx_path=tc_xlsx_path,
                ui_pdf_path=ui_pdf_path,
                output_dir=output_dir,
                form_path=TS_TEMPLATE_PATH,
            )
            payload["request_id"] = request_id
            if (
                    not payload.get("ok")
                    and payload.get("error")
                    and "요구사항 ID를 찾지 못했습니다" not in str(payload.get("error"))
            ):
                payload["error"] = (
                    f"{payload['error']}\n"
                    f"선택된 단위시험 케이스 파일: {tc_xlsx_name}\n"
                    f"선택된 사용자인터페이스설계서 파일: {ui_pdf_name}"
                )
            apply_download_stem(payload, uploaded_stem(template_xlsx_name, "template_xlsx"))
            attach_file_downloads(payload, delete_after_download=True, cleanup_root=temp_dir)
            preserve_temp_dir = bool(payload.get("download_files"))
            remove_runtime_path(template_xlsx_path)
            remove_runtime_path(tc_xlsx_path)
            remove_runtime_path(ui_pdf_path)

            log_event(
                "qa.ts.done",
                request_id=request_id,
                ok=payload.get("ok"),
                count=payload.get("count"),
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
