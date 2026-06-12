# 웹 실행 중 HWP/HWPX 변환을 준비하고 오래 걸리는 한글 프로세스를 제한합니다.
from __future__ import annotations

import csv
from pathlib import Path
import shutil
import subprocess
import sys
import time

from app_runtime import BASE_DIR
from document_update.hwp_convert import needs_hwp_to_hwpx_conversion
from document_update.hwpx_text import is_hwpx_zip


HWP_CONVERSION_TIMEOUT_SECONDS = 120
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def running_hwp_process_ids() -> set[int]:
    """한글 변환 타임아웃 때 새로 뜬 한글 프로세스만 정리하기 위해 현재 PID를 읽는다."""
    powershell_command = [
        "powershell",
        "-NoProfile",
        "-Command",
        "Get-Process -Name Hwp -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }",
    ]

    try:
        result = subprocess.run(
            powershell_command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=5,
            creationflags=CREATE_NO_WINDOW,
        )
        if result.returncode == 0:
            return {
                int(line.strip())
                for line in result.stdout.splitlines()
                if line.strip().isdigit()
            }
    except Exception:
        pass

    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=5,
            creationflags=CREATE_NO_WINDOW,
        )
    except Exception:
        return set()

    process_ids: set[int] = set()
    for row in csv.reader(result.stdout.splitlines()):
        if len(row) < 2 or row[0].lower() != "hwp.exe":
            continue
        try:
            process_ids.add(int(row[1]))
        except ValueError:
            continue
    return process_ids


def stop_new_hwp_processes(existing_process_ids: set[int]) -> None:
    """변환 작업이 멈췄을 때 이번 요청에서 새로 생긴 한글 프로세스만 종료한다."""
    for _ in range(3):
        process_ids = running_hwp_process_ids() - existing_process_ids
        if not process_ids:
            time.sleep(0.5)
            continue

        for process_id in process_ids:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"Stop-Process -Id {process_id} -Force -ErrorAction SilentlyContinue",
                ],
                capture_output=True,
                text=True,
                creationflags=CREATE_NO_WINDOW,
            )
            subprocess.run(
                ["taskkill", "/PID", str(process_id), "/F", "/T"],
                capture_output=True,
                text=True,
                creationflags=CREATE_NO_WINDOW,
            )

        time.sleep(0.5)


def convert_hwp_to_hwpx_with_timeout(input_path: Path, output_path: Path) -> Path:
    """한글 자동화 변환이 서버 요청을 오래 붙잡지 않도록 별도 프로세스에서 제한 시간만 기다린다."""
    existing_hwp_process_ids = running_hwp_process_ids()
    command = [
        sys.executable,
        "-X",
        "utf8",
        "-m",
        "document_update.hwp_convert",
        str(input_path),
        str(output_path),
    ]

    try:
        result = subprocess.run(
            command,
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=HWP_CONVERSION_TIMEOUT_SECONDS,
            creationflags=CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired as exc:
        stop_new_hwp_processes(existing_hwp_process_ids)
        raise RuntimeError(
            f"HWP를 HWPX로 변환하는 데 {HWP_CONVERSION_TIMEOUT_SECONDS}초가 넘게 걸렸습니다. "
            "HWPX로 저장한 파일을 업로드해 주세요."
        ) from exc

    if result.returncode != 0:
        stop_new_hwp_processes(existing_hwp_process_ids)
        message = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(message or "HWP를 HWPX로 변환하지 못했습니다.")

    if not output_path.exists():
        raise RuntimeError("HWPX 변환 결과 파일이 생성되지 않았습니다.")

    return output_path


def prepare_target_file(target_file: Path, temp_dir: Path) -> tuple[Path, bool]:
    """바이너리 한글 파일을 먼저 편집 가능한 압축 기반 한글 확장 파일로 변환한다."""
    if target_file.suffix.lower() == ".hwp" and is_hwpx_zip(target_file):
        converted_dir = temp_dir / "converted"
        converted_dir.mkdir(parents=True, exist_ok=True)
        converted_path = converted_dir / f"{target_file.stem}.hwpx"
        shutil.copy2(target_file, converted_path)
        return converted_path, True

    if not needs_hwp_to_hwpx_conversion(target_file):
        return target_file, False

    converted_dir = temp_dir / "converted"
    converted_path = converted_dir / f"{target_file.stem}.hwpx"
    return convert_hwp_to_hwpx_with_timeout(target_file, converted_path), True
