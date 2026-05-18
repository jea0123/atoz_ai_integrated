# 한글 자동화를 사용해 오래된 한글 파일을 편집 가능한 한글 확장 파일로 변환한다.
from __future__ import annotations

import argparse
from pathlib import Path
import sys

from app_runtime import log_event
from .hwpx_text import (
    clear_win32com_gen_cache,
    create_hwp_object,
    is_broken_win32com_cache_error,
    is_ole_hwp,
    is_hwpx_zip,
    open_hwp_document,
    win32,
)

try:
    import pythoncom
except ImportError:
    pythoncom = None


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
    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        # 일부 한글 설치 환경에서는 파일 경로 확인 모듈이 노출되지 않는다.
        pass


def _hide_hwp_window(hwp: object) -> None:
    try:
        hwp.XHwpWindows.Item(0).Visible = False
    except Exception:
        pass


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


def _convert_hwp_to_hwpx_once(input_path: Path, output_path: Path) -> None:
    hwp = None
    try:
        # 이미 떠 있는 한글 창에 붙으면 사용자의 문서나 이전 요청과 엮일 수 있으므로
        # 변환 요청마다 별도 한글 인스턴스를 만든다.
        log_event(
            "hwp_automation.convert",
            input_path=str(input_path),
            suffix=input_path.suffix.lower(),
            is_hwpx_zip=is_hwpx_zip(input_path),
            is_ole_hwp=is_ole_hwp(input_path),
        )
        hwp = create_hwp_object()
        _register_hwp_file_module(hwp)
        _hide_hwp_window(hwp)

        open_hwp_document(hwp, input_path)
        _save_as_hwpx(hwp, output_path)
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="HWP 파일을 HWPX로 변환")
    parser.add_argument("input_path", type=Path)
    parser.add_argument("output_path", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        convert_hwp_to_hwpx(args.input_path, args.output_path)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
