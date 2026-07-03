# 문서관리표준의 폴더명/문서명 표를 예상 폴더 경로 템플릿으로 변환합니다.
from __future__ import annotations

from pathlib import Path
import re

from .models import PathTemplate, StandardOutput
from .normalization import clean_text, normalize_for_match
from .standard_reader import extract_standard_text


FOLDER_ONLY_OUTPUT_NAMES: tuple[str, ...] = ()
RUN_PATTERN = re.compile(r"\S(?:.*?\S)?(?=\s{2,}\S|$)")
NUMBERED_NAME_PATTERN = re.compile(r"^\d{2}\.(?P<name>.+)$")
NOISE_PATTERN = re.compile(
    r"문서관리표준|MFDS-PP-|수입식품통합정보시스템|에이투지시스템|V\d+\.\d+|\d{4}\.\d{2}\.\d{2}"
)


def read_path_templates(
    standard_file: Path,
    outputs: list[StandardOutput],
    standard_text: str | None = None,
) -> list[PathTemplate]:
    # 표준 PDF에서 산출물별 예상 폴더 경로 템플릿을 읽는다.
    """문서관리표준의 폴더명/문서명 표에서 예상 경로 템플릿을 읽는다."""
    text = standard_text if standard_text is not None else extract_standard_text(standard_file)
    templates = parse_layout_path_templates(text)
    templates.extend(build_known_stage_templates(text))
    return deduplicate_templates(templates, outputs)


def parse_layout_path_templates(document_text: str) -> list[PathTemplate]:
    # PDF layout 텍스트의 폴더명/문서명 표를 들여쓰기 위치 기준으로 해석한다.
    templates: list[PathTemplate] = []
    levels: list[str | None] = [None, None, None]
    in_section = False

    for raw_line in document_text.splitlines():
        line = raw_line.rstrip()
        if "구분" in line and "폴더명/문서명" in line:
            in_section = True
            continue
        if not in_section:
            continue
        if NOISE_PATTERN.search(line):
            continue
        if not line.strip():
            continue

        runs = extract_runs(line)
        if not runs:
            continue

        leaf_names: list[str] = []
        for start, value in runs:
            if start < 6:
                if is_numbered(value):
                    levels[0] = value
                continue
            if start < 25:
                if is_numbered(value):
                    levels[1] = value
                continue
            if start < 44:
                if is_numbered(value):
                    levels[2] = value
                continue

            leaf_names.append(value)

        current_path = tuple(level for level in levels if level)
        for leaf_name in leaf_names:
            output_name = strip_number_prefix(leaf_name)
            if output_name and current_path:
                templates.append(
                    PathTemplate(
                        output_name=output_name,
                        template_path=current_path,
                        source=clean_text(line),
                    )
                )

    return templates


def extract_runs(line: str) -> list[tuple[int, str]]:
    # 한 줄 안에서 공백으로 분리된 텍스트 덩어리와 시작 위치를 뽑는다.
    runs: list[tuple[int, str]] = []
    for match in RUN_PATTERN.finditer(line):
        value = clean_text(match.group(0))
        if value:
            runs.append((match.start(), value))
    return runs


def build_known_stage_templates(document_text: str) -> list[PathTemplate]:
    # PDF 표가 깨지는 구간을 보정하기 위해 알려진 단계별 경로 템플릿을 추가한다.
    """행 병합 때문에 PDF 텍스트 순서가 뒤섞이는 시험/인도 핵심 경로를 보정한다."""
    compact = re.sub(r"\s+", "", document_text)
    templates: list[PathTemplate] = []

    if "05.시험" in compact:
        templates.extend(
            [
                make_template("통합시험결과서", "05.시험", "01.시험 & 전환", "01.통합시험"),
                make_template("성능시험결과서", "05.시험", "01.시험 & 전환", "02.시스템시험"),
                make_template("모의해킹수행결과서", "05.시험", "01.시험 & 전환", "02.시스템시험"),
                make_template("데이터전환결과서", "05.시험", "01.시험 & 전환", "03.데이터전환"),
                make_template("소스코드점검보고서", "05.시험", "02.시험단계점검", "01.취약점점검"),
                make_template("웹취약점진단결과서", "05.시험", "02.시험단계점검", "01.취약점점검"),
                make_template("데이터값진단계획서", "05.시험", "02.시험단계점검", "02.데이터값진단"),
                make_template("데이터값진단결과서", "05.시험", "02.시험단계점검", "02.데이터값진단"),
            ]
        )

    if "02.분석" in compact:
        templates.extend(
            [
                make_template("인터뷰계획서", "02.분석", "01.요구사항분석", "01.인터뷰계획서"),
                make_template("인터뷰결과서", "02.분석", "01.요구사항분석", "02.인터뷰결과서"),
                make_template("업무정의서", "02.분석", "02.현행시스템분석", "01.업무정의서"),
                make_template("현행아키텍처분석서", "02.분석", "02.현행시스템분석", "02.아키텍처분석"),
                make_template("총괄시험계획서", "02.분석", "03.분석단계시험계획", "01.총괄시험계획"),
                make_template("성능시험계획서", "02.분석", "03.분석단계시험계획", "02.성능(부하)시험계획"),
            ]
        )

    if "06.인도" in compact:
        templates.extend(
            [
                make_template("인수인계시험결과서", "06.인도", "01.인수인계", "01.인수인계"),
                make_template("운영자매뉴얼", "06.인도", "01.인수인계", "02.매뉴얼작성"),
                make_template("사용자매뉴얼", "06.인도", "01.인수인계", "02.매뉴얼작성"),
                make_template("교육계획서", "06.인도", "02.교육", "01.교육준비"),
                make_template("교육결과서", "06.인도", "02.교육", "02.교육수행"),
            ]
        )

    if "프로젝트종료" in compact or "03.프로젝트종료" in compact:
        templates.extend(
            [
                make_template("완료보고서", "01.프로젝트 시작", "03.프로젝트 종료"),
                make_template("사업완료보고서", "01.프로젝트 시작", "03.프로젝트 종료"),
                make_template("프로젝트완료보고서", "01.프로젝트 시작", "03.프로젝트 종료"),
            ]
        )

    if "05.설계단계시험계획" in compact:
        templates.extend(
            [
                make_template("단위시험케이스", "03.설계", "05.설계단계시험계획", "01.단위시험케이스"),
                make_template("통합시험시나리오", "03.설계", "05.설계단계시험계획", "02.통합시험시나리오"),
                make_template("인수인계시험시나리오", "03.설계", "05.설계단계시험계획", "03.인수인계시험시나리오"),
            ]
        )

    if "03.설계" in compact:
        templates.extend(
            [
                make_template("사용자인터페이스설계", "03.설계", "02.어플리케이션설계", "01.사용자인터페이스설계"),
                make_template("사용자인터페이스설계서", "03.설계", "02.어플리케이션설계", "01.사용자인터페이스설계"),
            ]
        )

    if "04.구현" in compact:
        templates.extend(
            [
                make_template("프로그램목록", "04.구현", "01.어플리케이션개발", "01.프로그램개발", "01.프로그램목록"),
                make_template("SQL튜닝요청서", "04.구현", "03.SQL개발"),
                make_template("SQL튜닝결과서", "04.구현", "03.SQL개발"),
                make_template("단위시험결과서", "04.구현", "02.단위시험", "01.단위시험"),
            ]
        )

    return templates


def make_template(output_name: str, *path_parts: str) -> PathTemplate:
    # 보정용 경로 템플릿 객체를 짧게 만들기 위한 헬퍼다.
    return PathTemplate(output_name, tuple(path_parts), "stage-correction")


def deduplicate_templates(
    templates: list[PathTemplate],
    outputs: list[StandardOutput],
) -> list[PathTemplate]:
    # 표준에 실제 존재하는 산출물 템플릿만 남기고 중복 경로를 정리한다.
    output_names = {
        normalize_for_match(name)
        for output in outputs
        for name in (output.output_name, *output.aliases)
        if name
    }
    by_name: dict[str, PathTemplate] = {}

    for template in templates:
        key = normalize_for_match(template.output_name)
        if not key or (key not in output_names and not is_folder_only_output_name(template.output_name)):
            continue

        current = by_name.get(key)
        if current is None or template.source == "stage-correction":
            by_name[key] = template

    return list(by_name.values())


def is_folder_only_output_name(output_name: str) -> bool:
    # 산출물 코드 표에는 없지만 폴더명/문서명 표에는 독립 문서로 있는 항목이다.
    key = normalize_for_match(output_name)
    return any(key == normalize_for_match(name) for name in FOLDER_ONLY_OUTPUT_NAMES)


def strip_number_prefix(value: str) -> str:
    # '01.산출물명' 같은 앞 번호를 제거한다.
    match = NUMBERED_NAME_PATTERN.match(clean_text(value))
    return clean_text(match.group("name")) if match else clean_text(value)


def is_numbered(value: str) -> bool:
    # 폴더 단계처럼 '01.xxx' 형식인지 확인한다.
    return bool(NUMBERED_NAME_PATTERN.match(value))
