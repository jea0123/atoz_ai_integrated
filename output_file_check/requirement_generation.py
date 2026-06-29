# 문서관리표준 반영이 끝난 기준 파일을 요구사항 파일명별로 복사 생성합니다.
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
import shutil
import unicodedata

from output_file_check.models import PathTemplate, StandardOutput
from output_file_check.normalization import filesystem_safe_stem, normalize_for_match, output_name_from_id


DEFAULT_REQUIREMENT_GENERATION_TARGETS: tuple[str, ...] = (
    "인터뷰결과서",
    "업무정의서",
    "개발데이터정의서",
    "개발데이터모델",
    "논리데이터모델",
    "논리데이터베이스설계서",
    "물리데이터모델",
    "물리데이터베이스설계서",
    "사용자인터페이스설계서",
    "단위시험케이스",
    "통합시험시나리오",
    "인수인계시험시나리오",
    "인수인계시험결과서",
    "프로그램목록",
    "단위시험결과서",
    "통합시험결과서",
)
APPLIED_OUTPUT_TARGET_SENTINELS = {
    "__applied__",
    "__matched__",
    "*",
    "all",
    "applied",
    "matched",
    "전체",
    "반영대상전체",
    "매칭전체",
}
APPLIED_OUTPUT_TARGET_LABEL = "반영 대상 전체"

REQUIREMENT_ID_PATTERN = re.compile(r"(?<![A-Z0-9])SFR-(?:[A-Z0-9]+-)*\d+(?![A-Z0-9])", re.IGNORECASE)
VERSION_PATTERN = re.compile(r"(?:^|[_\-\s])([vV]\d+(?:\.\d+)*)(?=$|[_\-\s.]|$)")
VERSION_TOKEN_REPLACE_PATTERN = re.compile(r"(^|[_\-\s])[vV]\d+(?:\.\d+)*", re.IGNORECASE)
ATTACHMENT_TAIL_PATTERN = re.compile(
    r"^[\s_-]*(?:[\[\(（［｛]\s*(?:별첨|첨부)\s*\d*[^)\]\}）］｝]*[\)\]\}）］｝]|(?:별첨|첨부)\s*\d+)",
    re.IGNORECASE,
)
DEFAULT_GENERATED_VERSION = "v0.1"
PROGRAM_SOURCE_TEMPLATE_PATH = (
    "04.구현",
    "01.어플리케이션개발",
    "01.프로그램개발",
    "02.프로그램소스",
)
PROGRAM_SOURCE_TAIL_PATH = ("01.프로그램개발", "02.프로그램소스")
HANDOVER_RESULT_KEY = normalize_for_match("인수인계시험결과서")
HANDOVER_CONFIRMATION_KEY = normalize_for_match("별첨1인수인계확인서")


@dataclass(frozen=True)
class RequirementSource:
    path: Path
    requirement_ids: tuple[str, ...]


@dataclass(frozen=True)
class OutputTemplate:
    output: StandardOutput
    template: PathTemplate | None
    source_paths: tuple[Path, ...]

    @property
    def source_path(self) -> Path:
        return self.source_paths[0]


@dataclass(frozen=True)
class RequirementGenerationResult:
    enabled: bool
    target_names: tuple[str, ...]
    target_count: int
    created_items: list[dict[str, object]]
    skipped_items: list[dict[str, object]]
    error_items: list[dict[str, object]]
    readme_path: Path | None = None
    removed_items: list[dict[str, object]] | None = None
    folder_items: list[dict[str, object]] | None = None


def generate_requirement_documents(
    dump_root: Path,
    outputs: list[StandardOutput],
    path_templates: list[PathTemplate],
    requirement_files: list[Path],
    standard_project_title: str,
    temp_dir: Path,
    fields: dict[str, str],
    *,
    apply_items: list[dict[str, object]] | None = None,
) -> RequirementGenerationResult:
    # 일반 요구사항 파일은 파일명에서, 제안요청서 입력은 임시 소스 파일명에서 ID를 읽는다.
    # 실제 복제 원본은 방금 문서관리표준 기준으로 정제된 산출물 파일이다.
    raw_target_names = fields.get("requirement_generation_targets")
    target_scope = parse_target_scope(raw_target_names)
    target_names = parse_target_names(raw_target_names, target_scope=target_scope)
    require_template_requirement_id = should_require_template_requirement_id(fields, target_scope)
    if not requirement_files:
        return RequirementGenerationResult(False, target_names, 0, [], [], [])

    sources, skipped_items = parse_requirement_sources(requirement_files)
    folder_items = (
        create_requirement_source_folders(dump_root, sources)
        if should_create_requirement_source_folders(fields)
        else []
    )
    targets = select_target_outputs(outputs, path_templates, target_names, target_scope=target_scope)
    output_templates, template_errors = select_output_templates(
        dump_root,
        targets,
        apply_items or [],
        report_missing=target_scope == "named",
        require_template_requirement_id=require_template_requirement_id,
    )
    error_items: list[dict[str, object]] = []
    created_items: list[dict[str, object]] = []

    if not targets:
        error_items.append(
            {
                "status": "error",
                "reason": "자동 생성 대상 산출물을 문서관리표준에서 찾지 못했습니다.",
                "target_names": ", ".join(target_names),
            }
        )
    error_items.extend(template_errors)

    if targets and not output_templates and not require_template_requirement_id:
        error_items.append(
            {
                "status": "error",
                "reason": "요구사항별 자동 생성에 사용할 정제 완료 기준 파일을 찾지 못했습니다.",
                "target_names": ", ".join(target_names),
            }
        )

    for source in sources:
        for requirement_id in source.requirement_ids:
            for output_template in output_templates:
                for template_path in output_template.source_paths:
                    created_items.append(
                        create_one_requirement_document(
                            source,
                            requirement_id,
                            output_template,
                            dump_root,
                            template_path,
                        )
                    )

    removed_items = remove_template_files_after_generation(
        dump_root,
        output_templates,
        created_items,
    )

    readme_path = write_requirement_generation_readme(
        dump_root,
        target_names,
        len(output_templates),
        created_items,
        skipped_items,
        error_items,
        removed_items,
        folder_items,
    )

    return RequirementGenerationResult(
        True,
        target_names,
        len(output_templates),
        created_items,
        skipped_items,
        error_items,
        readme_path,
        removed_items,
        folder_items,
    )


def parse_target_scope(raw_value: str | None) -> str:
    if not raw_value or not raw_value.strip():
        return "default"
    values = [item.strip() for item in re.split(r"[,;\n]+", raw_value) if item.strip()]
    if any(value.casefold() in APPLIED_OUTPUT_TARGET_SENTINELS for value in values):
        return "applied"
    return "named"


def parse_target_names(raw_value: str | None, *, target_scope: str | None = None) -> tuple[str, ...]:
    scope = target_scope or parse_target_scope(raw_value)
    if scope == "applied":
        return (APPLIED_OUTPUT_TARGET_LABEL,)
    if scope == "default":
        return DEFAULT_REQUIREMENT_GENERATION_TARGETS
    values = [item.strip() for item in re.split(r"[,;\n]+", raw_value) if item.strip()]
    return tuple(dict.fromkeys(values)) or DEFAULT_REQUIREMENT_GENERATION_TARGETS


def should_create_requirement_source_folders(fields: dict[str, str]) -> bool:
    raw = (
        fields.get("requirement_generation_create_source_folders")
        or fields.get("create_requirement_source_folders")
    )
    if raw is None or not str(raw).strip():
        return True
    return str(raw).strip().casefold() not in {"0", "false", "no", "n", "off"}


def should_require_template_requirement_id(fields: dict[str, str], target_scope: str) -> bool:
    raw = fields.get("requirement_generation_require_template_id")
    if raw is not None and str(raw).strip():
        return str(raw).strip().casefold() not in {"0", "false", "no", "n", "off"}
    return target_scope == "applied" or str(fields.get("artifact_category") or "").strip().casefold() == "management"


def parse_requirement_sources(requirement_files: list[Path]) -> tuple[list[RequirementSource], list[dict[str, object]]]:
    sources: list[RequirementSource] = []
    skipped: list[dict[str, object]] = []
    for path in requirement_files:
        requirement_ids = extract_requirement_ids(path.name)
        if not requirement_ids:
            skipped.append(
                {
                    "status": "skipped",
                    "source_file": path.name,
                    "reason": requirement_source_skip_reason(path),
                }
            )
            continue
        sources.append(
            RequirementSource(
                path=path,
                requirement_ids=requirement_ids,
            )
        )
    return sources, skipped


def extract_requirement_ids(filename: str) -> tuple[str, ...]:
    seen: set[str] = set()
    result: list[str] = []
    for match in REQUIREMENT_ID_PATTERN.finditer(filename):
        value = match.group(0).upper()
        if value in seen:
            continue
        result.append(value)
        seen.add(value)
    return tuple(result)


def requirement_source_skip_reason(path: Path) -> str:
    default_reason = "파일명에서 SFR로 시작하고 숫자 구간으로 끝나는 요구사항 ID를 찾지 못했습니다."
    try:
        if path.suffix.lower() != ".txt":
            return default_reason
        for line in path.read_text(encoding="utf-8", errors="ignore").splitlines()[:3]:
            text = line.strip()
            if text and not text.startswith("source:") and not text.startswith("requirement_id:"):
                return text
    except Exception:
        return default_reason
    return default_reason


def extract_template_version(path: Path) -> str:
    matches = VERSION_PATTERN.findall(path.stem)
    if not matches:
        return DEFAULT_GENERATED_VERSION
    return normalize_version(matches[-1])


def normalize_version(version: str) -> str:
    value = version.strip()
    if not value:
        return DEFAULT_GENERATED_VERSION
    if value[0] in {"v", "V"}:
        return f"v{value[1:]}"
    return f"v{value}"


def select_target_outputs(
    outputs: list[StandardOutput],
    path_templates: list[PathTemplate],
    target_names: tuple[str, ...],
    *,
    target_scope: str = "default",
) -> list[tuple[StandardOutput, PathTemplate | None]]:
    templates_by_name = {
        normalize_for_match(template.output_name): template
        for template in path_templates
    }
    selected: list[tuple[StandardOutput, PathTemplate | None]] = []
    seen: set[str] = set()
    for output in outputs:
        if target_scope != "applied" and not output_matches_targets(output, target_names):
            continue
        seen_key = output_match_key(output) if target_scope == "applied" else normalize_for_match(output.output_name)
        if seen_key in seen:
            continue
        selected.append((output, templates_by_name.get(normalize_for_match(output.output_name))))
        seen.add(seen_key)
    return selected


def output_matches_targets(output: StandardOutput, target_names: tuple[str, ...]) -> bool:
    output_keys = {
        key
        for name in (output.output_name, *output.aliases)
        for key in output_target_match_keys(name)
    }
    target_keys = {
        key
        for name in target_names
        for key in output_target_match_keys(name)
    }
    return bool(output_keys & target_keys)


def output_target_match_keys(name: str) -> tuple[str, ...]:
    key = normalize_for_match(name)
    if not key:
        return ()
    keys = [key]
    if key.endswith("서") and len(key) > 1:
        keys.append(key[:-1])
    return tuple(dict.fromkeys(keys))


def select_output_templates(
    dump_root: Path,
    targets: list[tuple[StandardOutput, PathTemplate | None]],
    apply_items: list[dict[str, object]],
    *,
    report_missing: bool,
    require_template_requirement_id: bool = False,
) -> tuple[list[OutputTemplate], list[dict[str, object]]]:
    selected: list[OutputTemplate] = []
    errors: list[dict[str, object]] = []
    seen_outputs: set[str] = set()

    for output, template in targets:
        output_key = output_match_key(output)
        if output_key in seen_outputs:
            continue
        seen_outputs.add(output_key)

        template_paths = select_applied_template_paths(
            output,
            apply_items,
            require_template_requirement_id=require_template_requirement_id,
        )
        if template_paths:
            selected.append(OutputTemplate(output, template, tuple(template_paths)))
            continue

        if report_missing:
            errors.append(
                {
                    "status": "error",
                    "output_id": output.output_id,
                    "output_name": output.output_name,
                    "target_path": str(resolve_target_folder(dump_root, output, template)),
                    "reason": "문서관리표준 반영이 완료된 기준 파일을 찾지 못했습니다.",
                }
            )

    return selected, errors


def select_applied_template_paths(
    output: StandardOutput,
    apply_items: list[dict[str, object]],
    *,
    require_template_requirement_id: bool = False,
) -> list[Path]:
    candidates: list[Path] = []
    seen: set[str] = set()
    for item in apply_items:
        if item.get("status") != "updated":
            continue
        if not apply_item_matches_output(item, output):
            continue
        raw_path = str(item.get("new_path") or "").strip()
        if not raw_path:
            continue
        path = Path(raw_path)
        path_key = str(path.resolve(strict=False)).casefold()
        if path_key in seen:
            continue
        if (
            path.exists()
            and is_supported_template_path(path)
            and (
                not require_template_requirement_id
                or template_has_requirement_id_tail(output, path)
            )
        ):
            candidates.append(path)
            seen.add(path_key)

    if not candidates:
        return []
    return select_distinct_template_variants(output, sorted(candidates, key=template_path_sort_key))


def template_has_requirement_id_tail(output: StandardOutput, path: Path) -> bool:
    tail = extract_template_tail(output, path.stem)
    return bool(tail and REQUIREMENT_ID_PATTERN.search(tail))


def select_distinct_template_variants(output: StandardOutput, candidates: list[Path]) -> list[Path]:
    selected: list[Path] = []
    seen: set[tuple[str, str, str]] = set()
    for path in candidates:
        key = template_variant_key(output, path)
        if key in seen:
            continue
        selected.append(path)
        seen.add(key)
    return selected


def template_variant_key(output: StandardOutput, path: Path) -> tuple[str, str, str]:
    suffix = path.suffix.lower()
    tail = extract_template_tail(output, path.stem) or ""
    if tail and (is_attachment_tail(tail) or is_handover_confirmation_template(output, path.stem)):
        return ("variant", suffix, normalize_template_variant_tail(tail))
    return ("base", suffix, "")


def normalize_template_variant_tail(tail: str) -> str:
    text = REQUIREMENT_ID_PATTERN.sub("SFR", tail)
    text = VERSION_TOKEN_REPLACE_PATTERN.sub(lambda match: f"{match.group(1)}v", text)
    return normalize_for_match(text)


def apply_item_matches_output(item: dict[str, object], output: StandardOutput) -> bool:
    item_output_name = normalize_for_match(str(item.get("output_name") or ""))
    output_names = [normalize_for_match(output.output_name), *(normalize_for_match(alias) for alias in output.aliases)]
    if item_output_name:
        return any(item_output_name == output_name for output_name in output_names if output_name)

    item_output_id = compact_identifier(str(item.get("output_id") or ""))
    output_id = compact_identifier(output.output_id)
    if item_output_id and output_id and item_output_id == output_id:
        return True

    return False


def output_match_key(output: StandardOutput) -> str:
    return f"{compact_identifier(output.output_id)}|{normalize_for_match(output.output_name)}"


def compact_identifier(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", value).casefold()


def is_supported_template_path(path: Path) -> bool:
    suffix = path.suffix.lower()
    return suffix in {
        ".hwp",
        ".hwpx",
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm",
        ".ppt",
        ".pptx",
        ".pptm",
        ".potx",
        ".potm",
        ".ppsx",
        ".ppsm",
    }


def template_path_sort_key(path: Path) -> tuple[int, str]:
    return (len(path.name), str(path).casefold())


def create_one_requirement_document(
    source: RequirementSource,
    requirement_id: str,
    output_template: OutputTemplate,
    dump_root: Path,
    template_path: Path | None = None,
) -> dict[str, object]:
    source_path = source.path
    output = output_template.output
    source_template_path = template_path or output_template.source_path
    target_folder = source_template_path.parent
    target_folder.mkdir(parents=True, exist_ok=True)

    try:
        suffix = source_template_path.suffix
        version = extract_template_version(source_template_path)
        expected_filename = build_generated_filename(
            output,
            requirement_id,
            version,
            suffix,
            template_path=source_template_path,
        )
        expected_path = target_folder / expected_filename
        if expected_path.exists():
            final_path = expected_path
            if same_path(expected_path, source_template_path):
                status = "exists"
            else:
                shutil.copy2(source_template_path, expected_path)
                status = "replaced"
        else:
            final_path = unique_file_path(expected_path)
            shutil.copy2(source_template_path, final_path)
            status = "created"

        return {
            "status": status,
            "source_file": source_path.name,
            "template_file": source_template_path.name,
            "requirement_id": requirement_id,
            "requirement_name": source_path.stem,
            "version": version,
            "output_id": output.output_id,
            "output_name": output.output_name,
            "target_path": str(final_path),
            "expected_filename": expected_filename,
            "warnings": [],
        }
    except Exception as exc:
        return {
            "status": "error",
            "source_file": source_path.name,
            "template_file": source_template_path.name,
            "requirement_id": requirement_id,
            "version": extract_template_version(source_template_path),
            "output_id": output.output_id,
            "output_name": output.output_name,
            "target_path": str(target_folder),
            "error": str(exc),
        }


def build_generated_filename(
    output: StandardOutput,
    requirement_id: str,
    version: str,
    suffix: str,
    *,
    template_path: Path | None = None,
) -> str:
    version_part = filesystem_safe_stem(normalize_version(version))
    if template_path:
        stem = build_generated_stem_from_template(output, requirement_id, version_part, template_path.stem)
    else:
        requirement_part = build_requirement_filename_part(requirement_id)
        stem = f"{build_standard_stem(output)}_{requirement_part}_{version_part}"
    return f"{filesystem_safe_stem(stem)}{suffix}"


def build_requirement_filename_part(requirement_id: str) -> str:
    return filesystem_safe_stem(requirement_id)


def build_generated_stem_from_template(
    output: StandardOutput,
    requirement_id: str,
    version: str,
    template_stem: str,
) -> str:
    base_stem = build_standard_stem(output)
    requirement_part = build_requirement_filename_part(requirement_id)
    tail = extract_template_tail(output, template_stem)
    if tail is None:
        return f"{base_stem}_{requirement_part}_{version}"

    if REQUIREMENT_ID_PATTERN.search(tail):
        tail = REQUIREMENT_ID_PATTERN.sub(requirement_part, tail)
    elif is_attachment_tail(tail):
        tail = f"_{requirement_part}_{version}{tail}"
    elif is_handover_confirmation_template(output, template_stem):
        return f"{base_stem}_{requirement_part}_{version}"
    else:
        return f"{base_stem}_{requirement_part}_{version}"

    tail = normalize_tail_versions(tail, version)
    if not VERSION_TOKEN_REPLACE_PATTERN.search(tail):
        tail = f"{tail}_{version}"
    return f"{base_stem}{tail}"


def is_handover_confirmation_template(output: StandardOutput, template_stem: str) -> bool:
    output_key = normalize_for_match(output.output_name)
    template_key = normalize_for_match(template_stem)
    return HANDOVER_RESULT_KEY in output_key and HANDOVER_CONFIRMATION_KEY in template_key


def is_attachment_tail(tail: str) -> bool:
    return bool(ATTACHMENT_TAIL_PATTERN.search(tail))


def extract_template_tail(output: StandardOutput, template_stem: str) -> str | None:
    prefixes = tuple(
        dict.fromkeys(
            value
            for value in (
                build_standard_stem(output),
                filesystem_safe_stem(build_standard_stem(output)),
                output.output_id,
                filesystem_safe_stem(output.output_id),
            )
            if value
        )
    )
    for prefix in prefixes:
        if len(template_stem) < len(prefix):
            continue
        if template_stem[:len(prefix)].casefold() != prefix.casefold():
            continue
        tail = template_stem[len(prefix):]
        if not tail or tail[0] in {"_", "-", "[", " "}:
            return tail
    return None


def normalize_tail_versions(tail: str, version: str) -> str:
    return VERSION_TOKEN_REPLACE_PATTERN.sub(lambda match: f"{match.group(1)}{version}", tail)


def same_path(left: Path, right: Path) -> bool:
    return str(left.resolve(strict=False)).casefold() == str(right.resolve(strict=False)).casefold()


def create_requirement_source_folders(
    dump_root: Path,
    sources: list[RequirementSource],
) -> list[dict[str, object]]:
    requirement_ids = unique_requirement_ids(sources)
    if not requirement_ids:
        return []

    source_root = resolve_program_source_root(dump_root)
    items = reset_program_source_root(dump_root, source_root)
    for requirement_id in requirement_ids:
        target_path = source_root / filesystem_safe_stem(requirement_id)
        item = {
            "requirement_id": requirement_id,
            "target_path": str(target_path),
        }
        existed = target_path.exists()
        try:
            target_path.mkdir(parents=True, exist_ok=True)
            items.append(
                {
                    **item,
                    "status": "exists" if existed else "created",
                    "folder_root": str(source_root),
                }
            )
        except Exception as exc:
            items.append({**item, "status": "error", "folder_root": str(source_root), "error": str(exc)})
    return items


def reset_program_source_root(dump_root: Path, source_root: Path) -> list[dict[str, object]]:
    if not is_path_inside(source_root, dump_root):
        return [
            {
                "status": "error",
                "folder_root": str(source_root),
                "target_path": str(source_root),
                "error": "결과 폴더 밖의 경로라 프로그램소스 폴더를 준비하지 않았습니다.",
            }
        ]

    if source_root.exists() and not source_root.is_dir():
        return [
            {
                "status": "error",
                "folder_root": str(source_root),
                "target_path": str(source_root),
                "error": "프로그램소스 경로가 폴더가 아니라서 준비하지 않았습니다.",
            }
        ]

    source_root.mkdir(parents=True, exist_ok=True)
    return []


def unique_requirement_ids(sources: list[RequirementSource]) -> tuple[str, ...]:
    seen: set[str] = set()
    result: list[str] = []
    for source in sources:
        for requirement_id in source.requirement_ids:
            key = requirement_id.casefold()
            if key in seen:
                continue
            result.append(requirement_id)
            seen.add(key)
    return tuple(result)


def resolve_program_source_root(dump_root: Path) -> Path:
    existing = find_program_source_root(dump_root)
    if existing:
        return existing
    return dump_root.joinpath(*PROGRAM_SOURCE_TEMPLATE_PATH)


def find_program_source_root(dump_root: Path) -> Path | None:
    expected = tuple(normalize_folder_part(part) for part in PROGRAM_SOURCE_TAIL_PATH)
    for path in sorted(iter_dirs(dump_root), key=lambda value: str(value).casefold()):
        try:
            parts = tuple(normalize_folder_part(part) for part in path.relative_to(dump_root).parts)
        except ValueError:
            continue
        if len(parts) >= len(expected) and parts[-len(expected):] == expected:
            return path
    return None


def normalize_folder_part(value: str) -> str:
    text = unicodedata.normalize("NFKC", value)
    text = re.sub(r"[^0-9A-Za-z가-힣]+", "", text)
    return text.casefold()


def remove_template_files_after_generation(
    dump_root: Path,
    output_templates: list[OutputTemplate],
    created_items: list[dict[str, object]],
) -> list[dict[str, object]]:
    successful_output_keys = {
        item_output_key(item)
        for item in created_items
        if item.get("status") != "error"
    }
    generated_paths = {
        str(Path(str(item.get("target_path") or "")).resolve(strict=False)).casefold()
        for item in created_items
        if item.get("status") != "error" and item.get("target_path")
    }
    removed_items: list[dict[str, object]] = []
    removed_seen: set[str] = set()

    for output_template in output_templates:
        output = output_template.output
        if output_match_key(output) not in successful_output_keys:
            continue

        removable_paths = [
            *output_template.source_paths,
            *stale_requirement_files_for_output(output_template, generated_paths),
        ]
        for source_path in removable_paths:
            path_key = str(source_path.resolve(strict=False)).casefold()
            if path_key in removed_seen or path_key in generated_paths:
                continue
            removed_seen.add(path_key)

            item = {
                "output_id": output.output_id,
                "output_name": output.output_name,
                "path": str(source_path),
            }
            if not is_path_inside(source_path, dump_root):
                removed_items.append(
                    {
                        **item,
                        "status": "error",
                        "error": "결과 폴더 밖의 파일이라 삭제하지 않았습니다.",
                    }
                )
                continue
            if not source_path.exists():
                removed_items.append({**item, "status": "missing"})
                continue
            if not source_path.is_file():
                removed_items.append({**item, "status": "skipped", "error": "파일이 아니라 삭제하지 않았습니다."})
                continue

            try:
                source_path.unlink()
                removed_items.append({**item, "status": "removed"})
            except Exception as exc:
                removed_items.append({**item, "status": "error", "error": str(exc)})

    return removed_items


def stale_requirement_files_for_output(
    output_template: OutputTemplate,
    keep_paths: set[str],
) -> list[Path]:
    paths: list[Path] = []
    seen: set[str] = set()
    folders = tuple(dict.fromkeys(path.parent for path in output_template.source_paths))
    for folder in folders:
        if not folder.exists() or not folder.is_dir():
            continue
        for path in folder.iterdir():
            path_key = str(path.resolve(strict=False)).casefold()
            if path_key in seen or path_key in keep_paths:
                continue
            if not is_stale_requirement_file(path, output_template.output):
                continue
            paths.append(path)
            seen.add(path_key)
    return paths


def is_stale_requirement_file(path: Path, output: StandardOutput) -> bool:
    if not path.is_file():
        return False
    if not is_supported_template_path(path):
        return False
    if not REQUIREMENT_ID_PATTERN.search(path.stem):
        return False

    stem_key = normalize_for_match(path.stem)
    output_keys = [
        normalize_for_match(output.output_id),
        normalize_for_match(output.output_name),
        *(normalize_for_match(alias) for alias in output.aliases),
    ]
    return any(output_key and output_key in stem_key for output_key in output_keys)


def is_path_inside(path: Path, root: Path) -> bool:
    try:
        path.resolve(strict=False).relative_to(root.resolve(strict=False))
        return True
    except ValueError:
        return False


def item_output_key(item: dict[str, object]) -> str:
    return f"{compact_identifier(str(item.get('output_id') or ''))}|{normalize_for_match(str(item.get('output_name') or ''))}"


def build_standard_stem(output: StandardOutput) -> str:
    id_name = output_name_from_id(output.output_id)
    if id_name and normalize_for_match(id_name) == normalize_for_match(output.output_name):
        return output.output_id
    if id_name:
        return output.output_id
    return f"{output.output_id}-{output.output_name}" if output.output_id else output.output_name


def resolve_target_folder(dump_root: Path, output: StandardOutput, template: PathTemplate | None) -> Path:
    if template:
        template_dir = find_template_dir(dump_root, template.template_path)
        if template_dir:
            if folder_name_matches_output(template_dir.name, output):
                return template_dir
            child = find_matching_child_dir(template_dir, output)
            if child:
                return child
            return template_dir / filesystem_safe_stem(output.output_name)

    existing = find_matching_dir_anywhere(dump_root, output)
    if existing:
        return existing

    if template:
        return dump_root.joinpath(*template.template_path) / filesystem_safe_stem(output.output_name)
    return dump_root / filesystem_safe_stem(output.output_name)


def find_template_dir(dump_root: Path, template_path: tuple[str, ...]) -> Path | None:
    direct = dump_root.joinpath(*template_path)
    if direct.exists() and direct.is_dir():
        return direct

    expected = tuple(normalize_for_match(part) for part in template_path)
    for path in iter_dirs(dump_root):
        try:
            parts = tuple(normalize_for_match(part) for part in path.relative_to(dump_root).parts)
        except ValueError:
            continue
        if endswith_parts(parts, expected) or contains_parts(parts, expected):
            return path
    return None


def find_matching_child_dir(parent: Path, output: StandardOutput) -> Path | None:
    for child in parent.iterdir():
        if child.is_dir() and folder_name_matches_output(child.name, output):
            return child
    return None


def find_matching_dir_anywhere(dump_root: Path, output: StandardOutput) -> Path | None:
    for path in iter_dirs(dump_root):
        if folder_name_matches_output(path.name, output):
            return path
    return None


def folder_name_matches_output(folder_name: str, output: StandardOutput) -> bool:
    folder_key = normalize_for_match(strip_folder_number_prefix(folder_name))
    output_keys = [normalize_for_match(output.output_name), *(normalize_for_match(alias) for alias in output.aliases)]
    return any(output_key and (folder_key == output_key or output_key in folder_key or folder_key in output_key) for output_key in output_keys)


def strip_folder_number_prefix(value: str) -> str:
    return re.sub(r"^\s*\d+[.)_-]?\s*", "", value)


def iter_dirs(root: Path):
    for path in root.rglob("*"):
        if path.is_dir():
            yield path


def endswith_parts(parts: tuple[str, ...], expected: tuple[str, ...]) -> bool:
    if not expected or len(parts) < len(expected):
        return False
    return parts[-len(expected):] == expected


def contains_parts(parts: tuple[str, ...], expected: tuple[str, ...]) -> bool:
    if not expected or len(parts) < len(expected):
        return False
    for index in range(len(parts) - len(expected) + 1):
        if parts[index:index + len(expected)] == expected:
            return True
    return False


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


def write_requirement_generation_readme(
    dump_root: Path,
    target_names: tuple[str, ...],
    target_count: int,
    created_items: list[dict[str, object]],
    skipped_items: list[dict[str, object]],
    error_items: list[dict[str, object]],
    removed_items: list[dict[str, object]],
    folder_items: list[dict[str, object]],
) -> Path:
    report_path = unique_file_path(dump_root / "README_요구사항별_자동생성.md")
    report_path.write_text(
        build_requirement_generation_readme(
            dump_root,
            target_names,
            target_count,
            created_items,
            skipped_items,
            error_items,
            removed_items,
            folder_items,
        ),
        encoding="utf-8",
    )
    return report_path


def build_requirement_generation_readme(
    dump_root: Path,
    target_names: tuple[str, ...],
    target_count: int,
    created_items: list[dict[str, object]],
    skipped_items: list[dict[str, object]],
    error_items: list[dict[str, object]],
    removed_items: list[dict[str, object]],
    folder_items: list[dict[str, object]],
) -> str:
    warning_items = [item for item in created_items if item.get("status") == "created_with_warning"]
    failed_items = [item for item in created_items if item.get("status") == "error"]
    removed_ok = [item for item in removed_items if item.get("status") == "removed"]
    removed_errors = [item for item in removed_items if item.get("status") == "error"]
    folder_ok = [item for item in folder_items if item.get("status") != "error"]
    folder_created = [item for item in folder_items if item.get("status") == "created"]
    folder_errors = [item for item in folder_items if item.get("status") == "error"]
    lines = [
        "# 요구사항별 자동 생성 README",
        "",
        "## 요약",
        "",
        "| 항목 | 값 |",
        "| --- | --- |",
        f"| 생성 시각 | {markdown_cell(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))} |",
        f"| 결과 폴더 | {markdown_cell(str(dump_root))} |",
        f"| 대상 산출물명 | {markdown_cell(', '.join(target_names))} |",
        f"| 정제 기준 파일을 찾은 대상 산출물 | {target_count}건 |",
        f"| 요구사항 ID 폴더 | {len(folder_ok)}건 (신규 {len(folder_created)}건) |",
        f"| 생성 파일 | {len([item for item in created_items if item.get('status') != 'error'])}건 |",
        f"| 기존 기준 파일 삭제 | {len(removed_ok)}건 |",
        f"| 생성 경고 | {len(warning_items)}건 |",
        f"| 생성 오류 | {len(failed_items) + len(error_items) + len(removed_errors) + len(folder_errors)}건 |",
        f"| 요구사항 ID 없음 | {len(skipped_items)}건 |",
        "",
        "## 요구사항 ID 폴더",
        "",
    ]
    append_requirement_folder_section(lines, folder_items, dump_root)
    lines.extend([
        "",
        "## 생성 파일",
        "",
    ])
    append_created_section(lines, created_items, dump_root)
    lines.extend(["", "## 기존 기준 파일 삭제", ""])
    append_removed_section(lines, removed_items, dump_root)
    lines.extend(["", "## 요구사항 ID 없음", ""])
    append_skipped_section(lines, skipped_items)
    lines.extend(["", "## 생성 오류", ""])
    append_generation_error_section(lines, failed_items, [*error_items, *removed_errors, *folder_errors], dump_root)
    lines.append("")
    return "\n".join(lines)


def append_requirement_folder_section(lines: list[str], items: list[dict[str, object]], dump_root: Path) -> None:
    if not items:
        lines.append("없음")
        return
    lines.extend(["| 상태 | 요구사항ID | 폴더 | 확인 내용 |", "| --- | --- | --- | --- |"])
    for item in items:
        status = item.get("status")
        label = "생성" if status == "created" else "기존" if status == "exists" else "오류"
        detail = item.get("error") or "프로그램소스 하위 요구사항 폴더"
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(label),
                    markdown_cell(item.get("requirement_id", "")),
                    markdown_cell(relative_report_path(str(item.get("target_path", "")), dump_root)),
                    markdown_cell(detail),
                ]
            )
            + " |"
        )


def append_created_section(lines: list[str], items: list[dict[str, object]], dump_root: Path) -> None:
    items = [item for item in items if item.get("status") != "error"]
    if not items:
        lines.append("없음")
        return
    lines.extend(
        [
            "| 상태 | 요구사항ID | 산출물명 | 파일 | 확인 내용 |",
            "| --- | --- | --- | --- | --- |",
        ]
    )
    for item in items:
        warnings = item.get("warnings") or []
        detail_items = [str(value) for value in warnings] if isinstance(warnings, list) and warnings else ["정상 생성"]
        if item.get("template_file"):
            detail_items.append(f"기준 파일: {item.get('template_file')}")
        detail = " / ".join(detail_items)
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(generated_file_status_label(str(item.get("status") or ""))),
                    markdown_cell(item.get("requirement_id", "")),
                    markdown_cell(item.get("output_name", "")),
                    markdown_cell(relative_report_path(str(item.get("target_path", "")), dump_root)),
                    markdown_cell(detail),
                ]
            )
            + " |"
        )


def generated_file_status_label(status: str) -> str:
    if status == "created_with_warning":
        return "경고"
    if status == "exists":
        return "기존"
    return "생성"


def append_removed_section(lines: list[str], items: list[dict[str, object]], dump_root: Path) -> None:
    visible_items = [item for item in items if item.get("status") in {"removed", "error"}]
    if not visible_items:
        lines.append("없음")
        return
    lines.extend(["| 상태 | 산출물명 | 파일 | 확인 내용 |", "| --- | --- | --- | --- |"])
    for item in visible_items:
        status = "삭제" if item.get("status") == "removed" else "오류"
        detail = item.get("error") or "요구사항 파일명 생성본으로 대체"
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(status),
                    markdown_cell(item.get("output_name", "")),
                    markdown_cell(relative_report_path(str(item.get("path", "")), dump_root)),
                    markdown_cell(detail),
                ]
            )
            + " |"
        )


def append_skipped_section(lines: list[str], items: list[dict[str, object]]) -> None:
    if not items:
        lines.append("없음")
        return
    lines.extend(["| 원본 파일 | 사유 |", "| --- | --- |"])
    for item in items:
        lines.append(
            "| "
            + " | ".join([markdown_cell(item.get("source_file", "")), markdown_cell(item.get("reason", ""))])
            + " |"
        )


def append_generation_error_section(
    lines: list[str],
    failed_items: list[dict[str, object]],
    error_items: list[dict[str, object]],
    dump_root: Path,
) -> None:
    rows = [*failed_items, *error_items]
    if not rows:
        lines.append("없음")
        return
    lines.extend(["| 요구사항ID | 산출물명 | 위치 | 오류 |", "| --- | --- | --- | --- |"])
    for item in rows:
        lines.append(
            "| "
            + " | ".join(
                [
                    markdown_cell(item.get("requirement_id", "")),
                    markdown_cell(item.get("output_name", "")),
                    markdown_cell(relative_report_path(str(item.get("target_path") or item.get("path") or ""), dump_root)),
                    markdown_cell(item.get("error") or item.get("reason") or ""),
                ]
            )
            + " |"
        )


def relative_report_path(path_text: str, dump_root: Path) -> str:
    if not path_text:
        return ""
    path = Path(path_text)
    try:
        return str(path.relative_to(dump_root))
    except ValueError:
        return path_text


def markdown_cell(value: object) -> str:
    text = str(value) if value is not None else ""
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    return text.replace("|", "\\|") or "-"
