# 표준 문서, 한글 문서, 엑셀 파일에서 인공지능과 규칙 매칭에 쓸 텍스트를 추출한다.
# HWP/HWPX/PDF/Office 파일에서 텍스트를 추출합니다. 폴더 매칭은 첫 장/표지 추출 함수를 씁니다.
from __future__ import annotations

from pathlib import Path
import importlib.util
import os
import re
import shutil
import struct
import sys
import tempfile
import winreg
import zipfile
import zlib

from app_runtime import log_event
from .patterns import TEXT_NODE_PATTERN

try:
    import win32com.client as win32
except ImportError:
    win32 = None

try:
    import pythoncom
except ImportError:
    pythoncom = None

try:
    import olefile
except ImportError:
    olefile = None


OOXML_TEXT_SUFFIXES = {
    ".xlsx", ".xlsm", ".xltx", ".xltm",
}
HWP_SECURITY_MODULE_NAME = "FilePathCheckerModule"
HWP_SECURITY_MODULE_DLL = "FilePathCheckerModule.dll"
HWP_SECURITY_REGISTRY_KEYS = (
    r"Software\HNC\HwpAutomation\Modules",
    r"Software\HNC\HwpCtrl\Modules",
)
LINESEG_ARRAY_PATTERN = re.compile(
    r"<(?P<tag>(?:\w+:)?linesegarray)\b[^>]*/>|"
    r"<(?P<tag2>(?:\w+:)?linesegarray)\b[^>]*>.*?</(?P=tag2)>",
    re.DOTALL | re.IGNORECASE,
)
def windows_long_path(path: Path) -> str:
    if os.name != "nt":
        return str(path)

    absolute = str(path.resolve(strict=False))
    if absolute.startswith("\\\\?\\"):
        return absolute
    if absolute.startswith("\\\\"):
        return "\\\\?\\UNC\\" + absolute.lstrip("\\")
    return "\\\\?\\" + absolute


def is_hwpx_zip(file_path: Path) -> bool:
    try:
        with zipfile.ZipFile(windows_long_path(file_path), "r") as zf:
            return "Contents/section0.xml" in zf.namelist()
    except (OSError, zipfile.BadZipFile):
        return False


def strip_hwpx_line_seg_arrays(xml: str) -> tuple[str, int]:
    """Remove stale HWPX layout cache after direct text edits."""
    return LINESEG_ARRAY_PATTERN.subn("", xml)


def is_zip_container(file_path: Path) -> bool:
    try:
        with zipfile.ZipFile(file_path, "r"):
            return True
    except (OSError, zipfile.BadZipFile):
        return False


def is_ole_hwp(file_path: Path) -> bool:
    return bool(olefile and olefile.isOleFile(str(file_path)))


def extract_text_from_hwpx(file_path: Path) -> str:
    """압축 기반 한글 확장 파일에서 텍스트를 추출한다."""
    texts: list[str] = []

    with zipfile.ZipFile(file_path, "r") as zf:
        for name in zf.namelist():
            if not name.lower().endswith(".xml"):
                continue

            try:
                xml_data = zf.read(name).decode("utf-8", errors="ignore")
                texts.extend(
                    match.group("body").strip()
                    for match in TEXT_NODE_PATTERN.finditer(xml_data)
                    if match.group("body").strip()
                )
            except Exception:
                continue

    return "\n".join(texts)


def extract_text_from_ooxml(file_path: Path) -> str:
    """엑셀 문서에서 보이는 문자열을 추출한다."""
    texts: list[str] = []

    with zipfile.ZipFile(file_path, "r") as zf:
        for name in zf.namelist():
            lower_name = name.lower()
            if not lower_name.endswith(".xml"):
                continue
            if lower_name.startswith("_rels/") or lower_name.endswith(".rels"):
                continue

            try:
                xml_data = zf.read(name).decode("utf-8", errors="ignore")
                texts.extend(
                    match.group("body").strip()
                    for match in TEXT_NODE_PATTERN.finditer(xml_data)
                    if match.group("body").strip()
                )
            except Exception:
                continue

    return "\n".join(texts)


def extract_text_nodes_limited(xml_data: str, max_chars: int) -> str:
    texts: list[str] = []
    total = 0
    for match in TEXT_NODE_PATTERN.finditer(xml_data):
        body = match.group("body").strip()
        if not body:
            continue
        texts.append(body)
        total += len(body) + 1
        if total >= max_chars:
            break
    return "\n".join(texts)[:max_chars]


def extract_text_from_hwpx_cover(file_path: Path, max_chars: int = 5000) -> str:
    with zipfile.ZipFile(file_path, "r") as zf:
        names = zf.namelist()
        by_lower = {name.lower(): name for name in names}
        preferred = by_lower.get("contents/section0.xml")
        section_names = sorted(
            (
                name
                for name in names
                if name.lower().startswith("contents/section") and name.lower().endswith(".xml")
            ),
            key=lambda name: hwp_section_sort_key([name]),
        )
        xml_names = [preferred] if preferred else []
        xml_names.extend(name for name in section_names if name not in xml_names)
        if not xml_names:
            xml_names = [name for name in names if name.lower().endswith(".xml")]

        texts: list[str] = []
        total = 0
        for name in xml_names:
            try:
                text = extract_text_nodes_limited(zf.read(name).decode("utf-8", errors="ignore"), max_chars)
            except Exception:
                continue
            if not text.strip():
                continue
            texts.append(text)
            total += len(text) + 1
            if total >= max_chars:
                break
        return "\n".join(texts)[:max_chars]


def extract_text_from_ooxml_cover(file_path: Path, max_chars: int = 5000) -> str:
    suffix = file_path.suffix.lower()
    with zipfile.ZipFile(file_path, "r") as zf:
        names = zf.namelist()
        by_lower = {name.lower(): name for name in names}
        if suffix in {".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"}:
            preferred_names = ["ppt/slides/slide1.xml"]
        elif suffix in {".docx", ".docm", ".dotx", ".dotm"}:
            preferred_names = ["word/document.xml"]
        elif suffix in OOXML_TEXT_SUFFIXES:
            preferred_names = ["xl/sharedstrings.xml", "xl/worksheets/sheet1.xml"]
        else:
            preferred_names = []

        xml_names = [by_lower[name] for name in preferred_names if name in by_lower]
        if not xml_names:
            xml_names = [
                name
                for name in names
                if name.lower().endswith(".xml")
                and not name.lower().startswith("_rels/")
                and not name.lower().endswith(".rels")
            ]

        texts: list[str] = []
        total = 0
        for name in xml_names:
            try:
                text = extract_text_nodes_limited(zf.read(name).decode("utf-8", errors="ignore"), max_chars)
            except Exception:
                continue
            if not text.strip():
                continue
            texts.append(text)
            total += len(text) + 1
            if total >= max_chars:
                break
        return "\n".join(texts)[:max_chars]


def hwp_section_sort_key(section_path: list[str]) -> int:
    match = re.search(r"\d+", section_path[-1])
    return int(match.group(0)) if match else 0


def decode_hwp_paragraph_text(payload: bytes) -> str:
    text = payload.decode("utf-16le", errors="ignore")
    return "".join(ch if (ch >= " " or ch in "\r\n\t") else "\n" for ch in text)


def extract_text_from_ole_hwp(file_path: Path) -> str:
    """확장자가 한글 확장 형식이어도 실제로는 오래된 한글 형식일 수 있는 파일에서 텍스트를 추출한다."""
    if olefile is None:
        raise RuntimeError("OLE 기반 HWP를 읽으려면 olefile 패키지가 필요합니다.")

    texts: list[str] = []

    with olefile.OleFileIO(str(file_path)) as ole:
        if ole.exists("PrvText"):
            preview_text = ole.openstream("PrvText").read().decode("utf-16le", errors="ignore")
            if preview_text.strip():
                texts.append(preview_text)

        if not ole.exists("FileHeader"):
            return "\n".join(texts)

        file_header = ole.openstream("FileHeader").read()
        compressed = len(file_header) > 36 and bool(file_header[36] & 1)

        sections = sorted(
            [
                path
                for path in ole.listdir()
                if len(path) == 2 and path[0] == "BodyText" and path[1].startswith("Section")
            ],
            key=hwp_section_sort_key,
        )

        for section in sections:
            data = ole.openstream(section).read()

            if compressed:
                data = zlib.decompress(data, -15)

            pos = 0
            while pos + 4 <= len(data):
                record_header = struct.unpack_from("<I", data, pos)[0]
                pos += 4

                tag_id = record_header & 0x3FF
                size = (record_header >> 20) & 0xFFF

                if size == 0xFFF:
                    if pos + 4 > len(data):
                        break
                    size = struct.unpack_from("<I", data, pos)[0]
                    pos += 4

                payload = data[pos:pos + size]
                pos += size

                if tag_id == 67:
                    paragraph_text = decode_hwp_paragraph_text(payload)
                    if paragraph_text.strip():
                        texts.append(paragraph_text)

    return "\n".join(texts)


def extract_text_from_ole_hwp_cover(file_path: Path, max_chars: int = 5000) -> str:
    if olefile is None:
        raise RuntimeError("OLE 湲곕컲 HWP瑜??쎌쑝?ㅻ㈃ olefile ?⑦궎吏媛 ?꾩슂?⑸땲??")

    texts: list[str] = []

    with olefile.OleFileIO(str(file_path)) as ole:
        if ole.exists("PrvText"):
            preview_text = ole.openstream("PrvText").read().decode("utf-16le", errors="ignore")
            if preview_text.strip():
                return preview_text[:max_chars]

        if not ole.exists("FileHeader"):
            return ""

        file_header = ole.openstream("FileHeader").read()
        compressed = len(file_header) > 36 and bool(file_header[36] & 1)

        sections = sorted(
            [
                path
                for path in ole.listdir()
                if len(path) == 2 and path[0] == "BodyText" and path[1].startswith("Section")
            ],
            key=hwp_section_sort_key,
        )
        if not sections:
            return ""

        data = ole.openstream(sections[0]).read()
        if compressed:
            data = zlib.decompress(data, -15)

        pos = 0
        total = 0
        while pos + 4 <= len(data):
            record_header = struct.unpack_from("<I", data, pos)[0]
            pos += 4

            tag_id = record_header & 0x3FF
            size = (record_header >> 20) & 0xFFF

            if size == 0xFFF:
                if pos + 4 > len(data):
                    break
                size = struct.unpack_from("<I", data, pos)[0]
                pos += 4

            payload = data[pos:pos + size]
            pos += size

            if tag_id != 67:
                continue

            paragraph_text = decode_hwp_paragraph_text(payload)
            if not paragraph_text.strip():
                continue

            texts.append(paragraph_text)
            total += len(paragraph_text) + 1
            if total >= max_chars:
                break

    return "\n".join(texts)[:max_chars]


def extract_text_from_hwp_cover(file_path: Path, max_chars: int = 5000) -> str:
    if is_ole_hwp(file_path):
        return extract_text_from_ole_hwp_cover(file_path, max_chars=max_chars)
    return extract_text_from_hwp(file_path)[:max_chars]


def extract_text_from_hwp(file_path: Path) -> str:
    if is_ole_hwp(file_path):
        return extract_text_from_ole_hwp(file_path)

    if win32 is None or pythoncom is None:
        raise RuntimeError(
            "이 HWP 파일은 현재 환경에서 텍스트를 추출할 수 없습니다. "
            "가능하면 HWPX 또는 XLSX 형식으로 저장해서 넣어주세요."
        )

    pythoncom.CoInitialize()
    hwp = None

    try:
        log_event(
            "hwp_automation.text_extract",
            file=str(file_path),
            suffix=file_path.suffix.lower(),
            is_hwpx_zip=is_hwpx_zip(file_path),
        )
        hwp = create_hwp_object()
        register_hwp_file_path_checker(hwp)
        open_hwp_document(hwp, file_path)
        return hwp.GetTextFile("TEXT", "")
    finally:
        if hwp is not None:
            hwp.Quit()
        pythoncom.CoUninitialize()


def is_broken_win32com_cache_error(exc: BaseException) -> bool:
    message = str(exc)
    return "CLSIDToClassMap" in message or "win32com.gen_py" in message


def clear_win32com_gen_cache() -> bool:
    """깨진 pywin32 COM 생성 캐시를 지운다. 캐시는 다음 실행 때 자동 재생성된다."""
    deleted = False
    candidates: list[Path] = []

    if win32 is not None:
        try:
            candidates.append(Path(win32.gencache.GetGeneratePath()))
        except Exception:
            pass

    candidates.append(Path(tempfile.gettempdir()) / "gen_py")

    temp_root = Path(tempfile.gettempdir()).resolve()
    seen: set[Path] = set()
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        if resolved in seen or not resolved.exists():
            continue
        seen.add(resolved)

        # 실수로 임시 폴더 밖을 지우지 않도록 pywin32 생성 캐시 위치만 허용한다.
        if resolved.name != "gen_py" and "gen_py" not in resolved.parts:
            continue
        if resolved != temp_root and temp_root not in resolved.parents:
            continue

        try:
            shutil.rmtree(resolved, ignore_errors=True)
            deleted = True
        except OSError:
            pass

    return deleted


def create_hwp_object() -> object:
    if win32 is None:
        raise RuntimeError("pywin32가 설치되어 있지 않아 한글 자동화 객체를 만들 수 없습니다.")

    try:
        return win32.DispatchEx("HWPFrame.HwpObject")
    except Exception as exc:
        if not is_broken_win32com_cache_error(exc):
            raise
        clear_win32com_gen_cache()
        return win32.DispatchEx("HWPFrame.HwpObject")


def find_hwp_security_module_path() -> Path | None:
    # pyhwpx에 포함된 한글 자동화 보안 승인 모듈 DLL을 찾는다.
    spec = importlib.util.find_spec("pyhwpx")
    if spec and spec.origin:
        candidate = Path(spec.origin).resolve().parent / HWP_SECURITY_MODULE_DLL
        if candidate.exists():
            return candidate

    for base in sys.path:
        candidate = Path(base) / "pyhwpx" / HWP_SECURITY_MODULE_DLL
        if candidate.exists():
            return candidate.resolve()

    return None


def ensure_hwp_security_module_registry() -> Path | None:
    # RegisterModule이 참조할 수 있도록 현재 사용자 레지스트리에 보안 모듈 DLL 경로를 등록한다.
    module_path = find_hwp_security_module_path()
    if module_path is None:
        log_event("hwp_security.module_missing", module=HWP_SECURITY_MODULE_NAME)
        return None

    for key_path in HWP_SECURITY_REGISTRY_KEYS:
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, HWP_SECURITY_MODULE_NAME, 0, winreg.REG_SZ, str(module_path))
        except OSError as exc:
            log_event(
                "hwp_security.registry_error",
                key=key_path,
                module=HWP_SECURITY_MODULE_NAME,
                error=str(exc),
            )

    return module_path


def register_hwp_file_path_checker(hwp: object) -> bool:
    # 한글 자동화의 외부 파일 접근 확인 팝업을 보안 모듈로 자동 승인한다.
    module_path = ensure_hwp_security_module_registry()
    if module_path is None:
        return False

    try:
        result = hwp.RegisterModule("FilePathCheckDLL", HWP_SECURITY_MODULE_NAME)
        log_event(
            "hwp_security.register",
            module=HWP_SECURITY_MODULE_NAME,
            dll=str(module_path),
            result=result,
        )
        return bool(result) or result is None
    except Exception as exc:
        log_event(
            "hwp_security.register_error",
            module=HWP_SECURITY_MODULE_NAME,
            dll=str(module_path),
            error=str(exc),
        )
        return False


def register_hwp_security_module(hwp: object) -> bool:
    return register_hwp_file_path_checker(hwp)


def open_hwp_document(hwp: object, file_path: Path) -> None:
    """한글 버전별로 다른 Open 인자 형식을 순서대로 시도한다."""
    path = str(file_path)
    if is_ole_hwp(file_path):
        attempts = (
            lambda: hwp.Open(path, "", ""),
            lambda: hwp.Open(path, "HWP", ""),
            lambda: hwp.Open(path),
            lambda: hwp.Open(path, "HWP", "forceopen:true"),
        )
    else:
        attempts = (
            lambda: hwp.Open(path, "", ""),
            lambda: hwp.Open(path, "HWP", ""),
            lambda: hwp.Open(path),
        )
    errors: list[str] = []

    for attempt in attempts:
        try:
            opened = attempt()
            if opened is False:
                errors.append("Open returned False")
                continue
            return
        except Exception as exc:
            errors.append(str(exc))

    detail = " / ".join(error for error in errors if error)
    raise RuntimeError(f"한글 프로그램이 파일을 열지 못했습니다: {file_path.name} ({detail})")


def add_bundled_site_packages() -> None:
    """기본 파이썬에 표준 문서 처리 패키지가 없을 때 번들 패키지 경로를 추가한다."""
    site_packages = (
        Path.home()
        / ".cache"
        / "codex-runtimes"
        / "codex-primary-runtime"
        / "dependencies"
        / "python"
        / "Lib"
        / "site-packages"
    )
    if site_packages.exists() and str(site_packages) not in sys.path:
        sys.path.append(str(site_packages))


def extract_text_from_pdf(file_path: Path) -> str:
    """표준 문서 처리 라이브러리로 문서관리표준에서 텍스트를 추출한다."""
    try:
        from pypdf import PdfReader
    except ImportError:
        add_bundled_site_packages()
        try:
            from pypdf import PdfReader
        except ImportError as exc:
            raise RuntimeError(
                "PDF 문서를 읽으려면 pypdf 패키지가 필요합니다. "
                "pip install pypdf 후 다시 실행하세요."
            ) from exc

    reader = PdfReader(str(file_path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def extract_text_from_pdf_cover(file_path: Path, max_chars: int = 5000) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        add_bundled_site_packages()
        try:
            from pypdf import PdfReader
        except ImportError as exc:
            raise RuntimeError(
                "PDF 臾몄꽌瑜??쎌쑝?ㅻ㈃ pypdf ?⑦궎吏媛 ?꾩슂?⑸땲?? "
                "pip install pypdf ???ㅼ떆 ?ㅽ뻾?섏꽭??"
            ) from exc

    reader = PdfReader(str(file_path))
    if not reader.pages:
        return ""
    return (reader.pages[0].extract_text() or "")[:max_chars]


# 파일 형식 추출
def extract_document_text(file_path: Path) -> str:
    """파일 확장자와 실제 컨테이너에 맞는 텍스트 추출기를 고른다."""
    suffix = file_path.suffix.lower()

    if suffix == ".pdf":
        return extract_text_from_pdf(file_path)

    if suffix in OOXML_TEXT_SUFFIXES:
        if not is_zip_container(file_path):
            raise RuntimeError(f"Office XML 형식이 아닙니다: {file_path.name}")
        return extract_text_from_ooxml(file_path)

    if is_hwpx_zip(file_path):
        return extract_text_from_hwpx(file_path)

    if suffix in {".hwp", ".hwpx"}:
        return extract_text_from_hwp(file_path)

    raise RuntimeError(
        "지원하지 않는 대상 파일 형식입니다. "
        "대상 파일은 HWP/HWPX 또는 XLSX/XLSM 계열을 사용해주세요."
    )
