# 모델 생성 호출 주소를 정리하고 프롬프트를 전송한다.
# Ollama generate API 호출을 한 곳에서 정리합니다. 프록시를 타지 않고 직접 요청합니다.
from __future__ import annotations

from urllib.parse import urlparse, urlunparse

import requests


OLLAMA_TIMEOUT_SECONDS = 20


def normalize_ollama_url(ollama_url: str) -> str:
    """요청을 보내기 전에 흔한 모델 호출 주소 형태를 정규화한다."""
    url = ollama_url.strip()

    if not url:
        return url

    if url.startswith(":"):
        url = f"http://127.0.0.1{url}"
    elif "://" not in url:
        url = f"http://{url}"

    parsed = urlparse(url)

    if parsed.scheme in {"http", "https"} and parsed.netloc.startswith(":"):
        parsed = parsed._replace(netloc=f"127.0.0.1{parsed.netloc}")

    if parsed.scheme in {"http", "https"} and not parsed.hostname:
        parsed = parsed._replace(netloc="127.0.0.1:11434")

    if parsed.path in {"", "/"}:
        parsed = parsed._replace(path="/api/generate")

    return urlunparse(parsed)


# AI 모델 생성 API를 호출
def generate(
    ollama_url: str,
    model: str,
    prompt: str,
    *,
    timeout: int | float | None = None,
    options: dict[str, object] | None = None,
    response_format: str | None = None,
) -> str:
    """로컬/사설망 모델 호출에서 시스템 프록시 설정을 무시하고 직접 요청한다."""
    session = requests.Session()
    session.trust_env = False
    payload: dict[str, object] = {"model": model, "prompt": prompt, "stream": False}
    if options:
        payload["options"] = options
    if response_format:
        payload["format"] = response_format

    response = session.post(
        normalize_ollama_url(ollama_url),
        json=payload,
        timeout=timeout or OLLAMA_TIMEOUT_SECONDS,
    )
    response.raise_for_status()

    data = response.json()
    return str(data.get("response", "")).strip()
