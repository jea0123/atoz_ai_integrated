# 한글 자동화를 사용해 오래된 한글 파일을 편집 가능한 한글 확장 파일로 변환한다.
from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import shutil
import sys
import threading
import time

from app_runtime import log_event
from .hwpx_text import (
    clear_win32com_gen_cache,
    create_hwp_object,
    is_broken_win32com_cache_error,
    is_ole_hwp,
    is_hwpx_zip,
    open_hwp_document,
    register_hwp_security_module,
    win32,
)

try:
    import pythoncom
except ImportError:
    pythoncom = None

try:
    import win32con
    import win32gui
except ImportError:
    win32con = None
    win32gui = None


def needs_hwp_to_hwpx_conversion(file_path: Path) -> bool:
    """바이너리 한글 파일이거나 압축 기반이 아닌 한글 확장 파일이면 변환이 필요하다고 판단한다."""
    suffix = file_path.suffix.lower()
    if is_hwpx_zip(file_path):
        return False
    if suffix == ".hwp":
        return True
    if suffix == ".hwpx" and not is_hwpx_zip(file_path):
        return True
    return False


def _register_hwp_file_module(hwp: object) -> None:
    register_hwp_security_module(hwp)


def _dialog_child_texts(window_handle: int) -> list[str]:
    if win32gui is None:
        return []

    texts: list[str] = []

    def collect_text(child_handle: int, _param: object) -> bool:
        text = win32gui.GetWindowText(child_handle).strip()
        if text:
            texts.append(text)
        return True

    try:
        win32gui.EnumChildWindows(window_handle, collect_text, None)
    except Exception:
        return texts
    return texts


def _normalized_dialog_text(value: str) -> str:
    return re.sub(r"[\s&()A-Za-z0-9:_-]+", "", value)


def _is_allow_all_text(value: str) -> bool:
    normalized = _normalized_dialog_text(value)
    return "모두허용" in normalized or ("모두" in normalized and "허용" in normalized)


def _is_allow_dialog(title: str, child_texts: list[str]) -> bool:
    if not any(_is_allow_all_text(text) for text in child_texts):
        return False

    joined = _normalized_dialog_text(" ".join([title, *child_texts]))
    return any(keyword in joined for keyword in ("한글", "보안", "허용", "접근", "파일"))


def _click_allow_all_dialogs_once() -> int:
    if win32gui is None or win32con is None:
        return 0

    clicked = 0

    def inspect_window(window_handle: int, _param: object) -> bool:
        nonlocal clicked
        try:
            title = win32gui.GetWindowText(window_handle).strip()
            child_texts = _dialog_child_texts(window_handle)
        except Exception:
            return True

        if not _is_allow_dialog(title, child_texts):
            return True

        def click_child(child_handle: int, _child_param: object) -> bool:
            nonlocal clicked
            try:
                child_text = win32gui.GetWindowText(child_handle).strip()
                if _is_allow_all_text(child_text):
                    win32gui.SendMessage(child_handle, win32con.BM_CLICK, 0, 0)
                    clicked += 1
                    log_event(
                        "hwp_automation.allow_all_clicked",
                        window_title=title,
                        button_text=child_text,
                    )
                    return False
            except Exception:
                return True
            return True

        try:
            win32gui.EnumChildWindows(window_handle, click_child, None)
        except Exception:
            return True
        return True

    try:
        win32gui.EnumWindows(inspect_window, None)
    except Exception:
        return clicked
    return clicked


def _start_allow_all_watcher() -> tuple[threading.Event, threading.Thread | None]:
    stop_event = threading.Event()
    if win32gui is None or win32con is None:
        log_event("hwp_automation.allow_all_watcher", started=False, reason="pywin32_gui_missing")
        return stop_event, None

    def watch() -> None:
        log_event("hwp_automation.allow_all_watcher", started=True)
        while not stop_event.is_set():
            _click_allow_all_dialogs_once()
            time.sleep(0.1)
        log_event("hwp_automation.allow_all_watcher", stopped=True)

    thread = threading.Thread(target=watch, daemon=True)
    thread.start()
    return stop_event, thread


def start_allow_all_watcher() -> tuple[threading.Event, threading.Thread | None]:
    """Start a lightweight desktop watcher that clicks HWP 'allow all' prompts."""
    return _start_allow_all_watcher()


def stop_allow_all_watcher(
    stop_event: threading.Event | None,
    thread: threading.Thread | None,
) -> None:
    if stop_event is not None:
        stop_event.set()
    if thread is not None:
        thread.join(timeout=1)


def _hide_hwp_window(hwp: object) -> None:
    try:
        hwp.XHwpWindows.Item(0).Visible = False
        log_event("hwp_automation.window_visible", visible=False)
    except Exception:
        log_event("hwp_automation.window_visible", visible="hide_failed")


def _save_as_hwpx(hwp: object, output_path: Path) -> bool:
    """한글 자동화에서 흔히 쓰이는 확장 문서 저장 방식을 차례로 시도한다."""
    errors: list[str] = []

    try:
        result = hwp.SaveAs(str(output_path), "HWPX", "")
        if output_path.exists():
            return bool(result) or True
    except Exception as exc:
        errors.append(f"SaveAs(path, HWPX, ''): {exc}")

    try:
        result = hwp.SaveAs(str(output_path), "HWPX")
        if output_path.exists():
            return bool(result) or True
    except Exception as exc:
        errors.append(f"SaveAs(path, HWPX): {exc}")

    try:
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(output_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        result = hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        if output_path.exists():
            return bool(result) or True
    except Exception as exc:
        errors.append(f"HAction FileSaveAs_S: {exc}")

    detail = " / ".join(errors)
    raise RuntimeError(
        "한글 프로그램으로 HWPX 저장에 실패했습니다. "
        "한글 버전이 HWPX 저장을 지원하는지 확인하세요."
        + (f" ({detail})" if detail else "")
    )


def _close_hwp_document(hwp: object) -> None:
    close_attempts = (
        lambda: hwp.Clear(1),
        lambda: hwp.Run("FileClose"),
    )
    for attempt in close_attempts:
        try:
            attempt()
            return
        except Exception:
            continue


def convert_hwp_to_hwpx(input_path: Path, output_path: Path) -> Path:
    """설치된 한글 프로그램으로 오래된 한글 파일을 압축 기반 한글 확장 파일로 변환한다."""
    if win32 is None or pythoncom is None:
        raise RuntimeError(
            "HWP를 HWPX로 변환하려면 pywin32와 한글 프로그램이 필요합니다. "
            "이 PC에서 python -m pip install pywin32 후 한글이 설치되어 있는지 확인하세요."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        output_path.unlink()

    # 웹 요청은 작업 스레드에서 처리되므로, 한글 자동화 객체를 만들기 전에
    # 각 스레드에서 자동화 환경을 초기화해야 한다.
    pythoncom.CoInitialize()

    try:
        try:
            _convert_hwp_to_hwpx_once(input_path, output_path)
        except Exception as exc:
            if not is_broken_win32com_cache_error(exc):
                raise
            clear_win32com_gen_cache()
            if output_path.exists():
                output_path.unlink()
            _convert_hwp_to_hwpx_once(input_path, output_path)
    finally:
        pythoncom.CoUninitialize()

    if not is_hwpx_zip(output_path):
        raise RuntimeError("HWPX로 변환했지만 ZIP 기반 HWPX 파일로 확인되지 않습니다.")

    return output_path


def convert_hwp_batch_to_hwpx(items: list[tuple[Path, Path]]) -> list[Path]:
    """Convert several HWP/OLE-HWPX files in one HWP automation session."""
    if win32 is None or pythoncom is None:
        raise RuntimeError(
            "HWP를 HWPX로 변환하려면 pywin32와 한글 프로그램이 필요합니다. "
            "이 PC에서 python -m pip install pywin32 후 한글이 설치되어 있는지 확인하세요."
        )

    pythoncom.CoInitialize()
    hwp = None
    allow_stop_event = None
    allow_thread = None
    converted_paths: list[Path] = []

    try:
        allow_stop_event, allow_thread = _start_allow_all_watcher()
        log_event("hwp_automation.batch_start", count=len(items))
        hwp = create_hwp_object()
        _register_hwp_file_module(hwp)
        _hide_hwp_window(hwp)

        for input_path, output_path in items:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            if output_path.exists():
                output_path.unlink()

            open_path = input_path
            temporary_open_path: Path | None = None
            try:
                if input_path.suffix.lower() == ".hwpx" and is_ole_hwp(input_path):
                    temporary_open_path = output_path.parent / f"{input_path.stem}.hwp"
                    shutil.copy2(input_path, temporary_open_path)
                    open_path = temporary_open_path
                    log_event(
                        "hwp_automation.open_alias",
                        original=str(input_path),
                        open_path=str(open_path),
                    )

                log_event(
                    "hwp_automation.batch_convert",
                    input_path=str(input_path),
                    suffix=input_path.suffix.lower(),
                    is_hwpx_zip=is_hwpx_zip(input_path),
                    is_ole_hwp=is_ole_hwp(input_path),
                )
                log_event("hwp_automation.convert_phase", phase="open_start")
                open_hwp_document(hwp, open_path)
                log_event("hwp_automation.convert_phase", phase="open_done")
                log_event("hwp_automation.convert_phase", phase="save_start")
                _save_as_hwpx(hwp, output_path)
                log_event("hwp_automation.convert_phase", phase="save_done")
                if not is_hwpx_zip(output_path):
                    raise RuntimeError(f"HWPX 변환 결과를 확인하지 못했습니다: {output_path.name}")
                converted_paths.append(output_path)
            finally:
                _close_hwp_document(hwp)
                if temporary_open_path is not None:
                    try:
                        temporary_open_path.unlink()
                    except OSError:
                        pass

        log_event("hwp_automation.batch_done", count=len(converted_paths))
        return converted_paths
    finally:
        stop_allow_all_watcher(allow_stop_event, allow_thread)
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _convert_hwp_to_hwpx_once(input_path: Path, output_path: Path) -> None:
    hwp = None
    open_path = input_path
    temporary_open_path: Path | None = None
    allow_stop_event: threading.Event | None = None
    allow_thread: threading.Thread | None = None
    try:
        allow_stop_event, allow_thread = _start_allow_all_watcher()
        if input_path.suffix.lower() == ".hwpx" and is_ole_hwp(input_path):
            temporary_open_path = output_path.parent / f"{input_path.stem}.hwp"
            shutil.copy2(input_path, temporary_open_path)
            open_path = temporary_open_path
            log_event(
                "hwp_automation.open_alias",
                original=str(input_path),
                open_path=str(open_path),
            )

        # 이미 떠 있는 한글 창에 붙으면 사용자의 문서나 이전 요청과 엮일 수 있으므로
        # 변환 요청마다 별도 한글 인스턴스를 만든다.
        log_event(
            "hwp_automation.convert",
            input_path=str(input_path),
            suffix=input_path.suffix.lower(),
            is_hwpx_zip=is_hwpx_zip(input_path),
            is_ole_hwp=is_ole_hwp(input_path),
        )
        log_event("hwp_automation.convert_phase", phase="create_start")
        hwp = create_hwp_object()
        log_event("hwp_automation.convert_phase", phase="create_done")
        _register_hwp_file_module(hwp)
        log_event("hwp_automation.convert_phase", phase="register_done")
        _hide_hwp_window(hwp)

        log_event("hwp_automation.convert_phase", phase="open_start")
        open_hwp_document(hwp, open_path)
        log_event("hwp_automation.convert_phase", phase="open_done")
        log_event("hwp_automation.convert_phase", phase="save_start")
        _save_as_hwpx(hwp, output_path)
        log_event("hwp_automation.convert_phase", phase="save_done")
    finally:
        stop_allow_all_watcher(allow_stop_event, allow_thread)
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass
        if temporary_open_path is not None:
            try:
                temporary_open_path.unlink()
            except OSError:
                pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="HWP 파일을 HWPX로 변환")
    parser.add_argument("input_path", type=Path, nargs="?")
    parser.add_argument("output_path", type=Path, nargs="?")
    parser.add_argument("--manifest", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        if args.manifest:
            manifest_items = json.loads(args.manifest.read_text(encoding="utf-8"))
            convert_hwp_batch_to_hwpx(
                [
                    (Path(item["input_path"]), Path(item["output_path"]))
                    for item in manifest_items
                ]
            )
        else:
            if args.input_path is None or args.output_path is None:
                raise RuntimeError("input_path와 output_path가 필요합니다.")
            convert_hwp_to_hwpx(args.input_path, args.output_path)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
