# 실행 경로, 로그, .env, AI/규칙 모드 판단처럼 여러 진입점이 공유하는 런타임 도우미입니다.
from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import re
import shutil


BASE_DIR = Path(__file__).resolve().parent

WEB_DIR = BASE_DIR / "web"
TEMPLATE_DIR = BASE_DIR / "templates"

WORK_DIR = BASE_DIR / "web_runtime"
TEMP_DIR = WORK_DIR / "temp"
RESULT_DIR = WORK_DIR / "results"

TS_TEMPLATE_PATH = TEMPLATE_DIR / "scenario_sheet_form.xlsx"
RESULT_TEMPLATE_PATH = TEMPLATE_DIR / "result_sheet_form.xlsx"

LOG_PATH = WORK_DIR / "web_app.log"
RUNTIME_ENV_PATH = BASE_DIR / ".env"
SENSITIVE_URL_PATTERN = re.compile(r"https?://[^\s'\"<>]+")
DEFAULT_MODEL = "exaone3.5:7.8b"


def log_event(event: str, **values: object) -> None:
    # 서버/매칭 진행 상황을 web_runtime/web_app.log에 한 줄로 남긴다.
    """디버그 로그를 서버 터미널과 로그 파일에 함께 남긴다."""
    WORK_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    extras = " ".join(f"{key}={sanitize_log_value(value)!r}" for key, value in values.items())
    line = f"[{timestamp}] {event}"
    if extras:
        line = f"{line} {extras}"

    write_log_line(line)


def log_message(message: str) -> None:
    """사람이 읽기 쉬운 진행 로그를 서버 콘솔과 로그 파일에 함께 남긴다."""
    WORK_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    write_log_line(f"[{timestamp}] {message}")


def write_log_line(line: str) -> None:
    print(line, flush=True)
    try:
        with LOG_PATH.open("a", encoding="utf-8") as log_file:
            log_file.write(line + "\n")
    except OSError:
        pass


def sanitize_log_value(value: object) -> object:
    # 로그에 외부 URL이 그대로 찍히지 않도록 민감한 URL 문자열을 숨긴다.
    if not isinstance(value, str):
        return value
    return SENSITIVE_URL_PATTERN.sub("[hidden-url]", value)


def parse_json_object(raw_text: str) -> dict[str, object]:
    # AI 응답에서 JSON 객체 부분만 찾아 dict로 파싱한다.
    text = raw_text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

    match = re.search(r"\{.*\}", text, re.DOTALL)
    if not match:
        return {}

    try:
        value = json.loads(match.group(0))
    except json.JSONDecodeError:
        return {}

    return value if isinstance(value, dict) else {}


def ensure_runtime_dirs() -> None:
    # 서버 실행에 필요한 런타임 폴더를 만든다.
    """서버 시작 시 작업 폴더와 결과 폴더를 만든다."""
    for path in (
        WORK_DIR,
        TEMP_DIR,
        RESULT_DIR,
    ):
        path.mkdir(parents=True, exist_ok=True)


def cleanup_runtime() -> None:
    # 서버 시작 시 중단된 요청 때문에 남은 임시 요청 폴더를 정리한다.
    """중단된 요청 때문에 남은 임시 요청 폴더를 정리한다."""
    if not WORK_DIR.exists():
        return

    for root, patterns in (
        (WORK_DIR, ("doc-update-*", "folder-check-*", "folder-apply-*", "c-*", "a-*")),
        (TEMP_DIR, ("qa-tc-*", "qa-ts-*")),
    ):
        if not root.exists():
            continue

        for pattern in patterns:
            for path in root.glob(pattern):
                remove_runtime_path(path)


def remove_runtime_path(path: Path) -> None:
    # 임시 파일/폴더를 지우되 실패해도 서버 흐름을 막지 않는다.
    """Windows 파일 잠금 때문에 정리에 실패해도 실제 요청 처리는 막지 않는다."""
    try:
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        elif path.is_file():
            path.unlink(missing_ok=True)
    except OSError:
        pass


def read_runtime_env() -> dict[str, str]:
    # .env 파일에서 OLLAMA_URL, MODEL 같은 실행 설정을 읽는다.
    """셸 환경변수나 다른 폴더가 아니라 현재 앱의 환경 파일만 읽는다."""
    if not RUNTIME_ENV_PATH.exists():
        return {}

    values: dict[str, str] = {}
    for raw_line in RUNTIME_ENV_PATH.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip().lstrip("\ufeff")
        value = value.strip().strip("\"'")
        if key:
            values[key] = value

    return values

def resolve_runtime_ollama_base_url(env: dict[str, str]) -> str | None:
    # .env에서 OLLAMA_BASE_URL을 읽고 끝에 /api/generate 또는 /api/chat이 붙은 경우 제거한다.
    raw = env.get("OLLAMA_BASE_URL") or env.get("OLLAMA_URL")
    if not raw:
        return None
    
    raw = raw.rstrip("/")
    for suffix in ("/api/generate", "/api/chat"):
        if raw.endswith(suffix):
            return raw[:-len(suffix)]
        
    return raw

def resolve_runtime_ollama_generate_url(env: dict[str, str]) -> str | None:
    # Ollama 생성 호출 URL을 만든다.
    base_url = resolve_runtime_ollama_base_url(env)
    if not base_url:
        return None
    return f"{base_url}/api/generate"

def resolve_runtime_ollama_chat_url(env: dict[str, str]) -> str | None:
    # Ollama 채팅 호출 URL을 만든다.
    base_url = resolve_runtime_ollama_base_url(env)
    if not base_url:
        return None
    return f"{base_url}/api/chat"


def resolve_runtime_model(env: dict[str, str]) -> str:
    # .env 또는 환경 변수에서 모델명을 찾고 없으면 기본 모델을 쓴다.
    return env.get("OLLAMA_MODEL") or DEFAULT_MODEL


def runtime_mode_payload() -> dict[str, object]:
    # 화면에 표시할 현재 AI/규칙 모드 상태를 만든다.
    env = read_runtime_env()
    ollama_url = resolve_runtime_ollama_base_url(env)
    model = resolve_runtime_model(env)
    if ollama_url:
        return {
            "mode": "rule_ai_fallback",
            "label": "규칙 매칭 · 미매칭 AI 확인",
            "model": model,
            "ollama_configured": True,
        }

    return {
        "mode": "rule_fallback_no_ollama",
        "label": "규칙 매칭만",
        "model": model,
        "ollama_configured": False,
    }


def selected_match_mode(fields: dict[str, str], ollama_url: str | None) -> tuple[str, str]:
    # 기본은 규칙 매칭을 먼저 쓰고, 비어 있는 산출물만 AI로 보강한다.
    requested = (fields.get("match_mode") or "rule_ai_fallback").strip()

    if requested == "rule_only":
        return "rule_only", "rule_only"
    if requested == "ai_first" and ollama_url:
        return "ai_first", "ai_first"
    if ollama_url:
        return "rule_ai_fallback", "rule_ai_fallback"
    return "rule_only", "rule_fallback_no_ollama"
