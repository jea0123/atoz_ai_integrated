from pathlib import Path
import shutil
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from uuid import uuid4

from output_file_check.folder_apply_ops import (
    apply_dumped_folder,
    apply_batch_candidate,
    apply_project_title_only_file,
    apply_unmatched_project_title_updates,
    build_apply_readme,
    build_cover_status,
    build_target_filename,
    cleanup_original_after_replacement,
    filter_apply_scope_candidates,
    project_title_for_update,
    validate_output_id,
)
from output_file_check.models import FileIdentity, MatchCandidate, ScannedFile, StandardOutput
from output_file_check.requirement_generation import RequirementGenerationResult


class FolderApplyOpsTest(unittest.TestCase):
    def test_validate_output_id_allows_project_specific_values(self) -> None:
        for output_id in ("11", "112233", "CUSTOM-ID", "MFDS-업무ID-YYMMDD-회의명"):
            with self.subTest(output_id=output_id):
                validate_output_id(output_id)

    def test_target_filename_uses_arbitrary_output_id_and_output_name(self) -> None:
        output = StandardOutput(
            output_id="112233",
            output_name="사업수행계획서",
        )

        filename = build_target_filename(output, Path("MFDS-PP-01-사업수행계획서_v1.0.hwp"))

        self.assertEqual("112233-사업수행계획서_v0.1.hwp", filename)

    def test_target_filename_preserves_sfr_tail_for_requirement_templates(self) -> None:
        output = StandardOutput(
            output_id="14",
            output_name="요구사항정의서",
        )

        filename = build_target_filename(output, Path("MFDS-PMC-01-요구사항정의서_SFR-ESS-001_v1.0.hwpx"))

        self.assertEqual("14-요구사항정의서_SFR-ESS-001_v0.1.hwpx", filename)

    def test_management_apply_scope_ignores_top_level_files(self) -> None:
        root = Path.cwd() / "dump"
        output = StandardOutput(output_id="20", output_name="risk")
        root_candidate = MatchCandidate(
            output=output,
            file=ScannedFile(root / "proposal.hwp", FileIdentity(project_title="project")),
            score=1.0,
            reason="test",
        )
        nested_candidate = MatchCandidate(
            output=output,
            file=ScannedFile(root / "01.project" / "risk.xlsx", FileIdentity(project_title="project")),
            score=1.0,
            reason="test",
        )

        filtered = filter_apply_scope_candidates(
            [root_candidate, nested_candidate],
            root,
            {"artifact_category": "management"},
        )

        self.assertEqual([nested_candidate], filtered)

    def test_apply_candidate_renames_file_when_cover_update_fails(self) -> None:
        source = Path.cwd() / "MFDS-OLD-99-사업수행계획서_v1.0.xlsx"
        output = StandardOutput(
            output_id="MFDS-PP-01-사업수행계획서",
            output_name="사업수행계획서",
        )
        candidate = MatchCandidate(
            output=output,
            file=ScannedFile(source, FileIdentity(project_title="기존 프로젝트")),
            score=1.0,
            reason="test",
        )

        with (
            patch("output_file_check.folder_apply_ops.prepare_target_file", return_value=(source, False)),
            patch(
                "output_file_check.folder_apply_ops.write_updated_document",
                side_effect=RuntimeError("표지 형식 미지원"),
            ),
            patch.object(Path, "replace", return_value=None),
            patch(
                "output_file_check.folder_apply_ops.apply_initial_revision_metadata",
                return_value={
                    "status": "skipped",
                    "revision_date": "2026-00-00",
                    "author": "",
                    "approval_author": "",
                    "cover_update_count": 0,
                    "revision_history_update_count": 0,
                    "error": "",
                },
            ),
        ):
            result = apply_batch_candidate(
                candidate,
                "신규 프로젝트",
                Path.cwd() / "temp",
                revision_metadata={
                    "revision_date": "2026-00-00",
                    "author": "",
                    "approval_author": "",
                },
                rename_files=True,
            )

        self.assertEqual("updated", result["status"])
        self.assertEqual("failed", result["cover_update_status"])
        self.assertIn("표지 형식 미지원", result["cover_update_error"])
        self.assertEqual("MFDS-PP-01-사업수행계획서_v0.1.xlsx", Path(str(result["new_path"])).name)

    def test_cleanup_original_after_replacement_reports_permission_error(self) -> None:
        source = Path.cwd() / "source.xlsx"
        target = Path.cwd() / "target.xlsx"

        with (
            patch.object(Path, "exists", return_value=True),
            patch.object(Path, "unlink", side_effect=PermissionError("locked")),
        ):
            error = cleanup_original_after_replacement(source, target)

        self.assertIn("locked", error)

    def test_project_title_only_update_keeps_unmatched_filename(self) -> None:
        source = Path.cwd() / "기술적용결과표.hwpx"
        converted = Path.cwd() / "기술적용결과표.hwpx"
        updated = Path.cwd() / "temp-updated.hwpx"

        with (
            patch("output_file_check.folder_apply_ops.prepare_target_file", return_value=(converted, False)),
            patch("output_file_check.folder_apply_ops.write_updated_project_title", return_value=(1, updated)),
            patch("output_file_check.folder_apply_ops.replace_file_with_fallback") as replace_mock,
            patch(
                "output_file_check.folder_apply_ops.apply_initial_revision_metadata",
                return_value={
                    "status": "updated",
                    "revision_date": "2026-00-00",
                    "author": "author",
                    "approval_author": "approval",
                    "cover_update_count": 3,
                    "revision_history_update_count": 0,
                    "error": "",
                },
            ) as metadata_mock,
            patch.object(Path, "exists", return_value=False),
        ):
            result = apply_project_title_only_file(
                source,
                "2025년도 수입식품통합정보시스템 고도화",
                "2026년도 수입식품통합정보시스템 고도화",
                Path.cwd() / "temp",
                {
                    "revision_date": "2026-00-00",
                    "author": "author",
                    "approval_author": "approval",
                },
            )

        self.assertEqual("updated", result["status"])
        self.assertTrue(result["project_only"])
        self.assertEqual("표준 외 문서", result["output_name"])
        self.assertEqual(1, result["cover_project_replace_count"])
        self.assertEqual(3, result["initial_revision_cover_update_count"])
        self.assertEqual("기술적용결과표.hwpx", Path(str(result["new_path"])).name)
        replace_mock.assert_called_once_with(updated, source)
        metadata_mock.assert_called_once()

    def test_project_title_only_update_scans_review_template_folder(self) -> None:
        dump_root = Path.cwd() / ".test-artifacts" / uuid4().hex
        source = dump_root / "양식" / "MFDS-PMC-02-요구사항추적표.xlsx"
        updated = dump_root / "updated.xlsx"
        source.parent.mkdir(parents=True, exist_ok=True)
        source.write_text("placeholder", encoding="utf-8")
        try:
            with (
                patch(
                    "output_file_check.folder_apply_ops.scan_folder",
                    return_value=[
                        ScannedFile(
                            source,
                            FileIdentity(project_title="2024년 지능형 수입식품 통합시스템 고도화"),
                        )
                    ],
                ),
                patch("output_file_check.folder_apply_ops.prepare_target_file", return_value=(source, False)),
                patch("output_file_check.folder_apply_ops.write_updated_project_title", return_value=(1, updated)),
                patch("output_file_check.folder_apply_ops.replace_file_with_fallback"),
                patch(
                    "output_file_check.folder_apply_ops.apply_initial_revision_metadata",
                    return_value={
                        "status": "updated",
                        "revision_date": "2026-00-00",
                        "author": "author",
                        "approval_author": "approval",
                        "cover_update_count": 3,
                        "revision_history_update_count": 0,
                        "error": "",
                    },
                ),
            ):
                items = apply_unmatched_project_title_updates(
                    dump_root,
                    [],
                    [],
                    [],
                    "2026년도 수입식품통합정보시스템 고도화",
                    dump_root / "temp",
                    set(),
                    revision_metadata={
                        "revision_date": "2026-00-00",
                        "author": "송아름",
                        "approval_author": "임채현",
                    },
                )

            self.assertEqual(1, len(items))
            self.assertEqual("updated", items[0]["status"])
            self.assertTrue(items[0]["project_only"])
        finally:
            shutil.rmtree(dump_root, ignore_errors=True)

    def test_unmatched_review_template_with_output_match_gets_full_apply(self) -> None:
        dump_root = Path.cwd() / ".test-artifacts" / uuid4().hex
        source = dump_root / "양식" / "MFDS-PMC-02-요구사항추적표(검사기준포함)_v1.0.xlsx"
        source.parent.mkdir(parents=True, exist_ok=True)
        source.write_text("placeholder", encoding="utf-8")
        output = StandardOutput("15", "요구사항추적표(검사기준포함)")
        try:
            with (
                patch(
                    "output_file_check.folder_apply_ops.scan_folder",
                    return_value=[
                        ScannedFile(
                            source,
                            FileIdentity(
                                project_title="2024년 지능형 수입식품 통합시스템 고도화",
                                document_title="요구사항추적표",
                                preview_text="2024년 지능형 수입식품 통합시스템 고도화 요구사항추적표",
                            ),
                        )
                    ],
                ),
                patch(
                    "output_file_check.folder_apply_ops.apply_batch_candidate",
                    return_value={"status": "updated", "project_only": False},
                ) as apply_mock,
                patch("output_file_check.folder_apply_ops.write_updated_project_title") as project_only_mock,
            ):
                items = apply_unmatched_project_title_updates(
                    dump_root,
                    [output],
                    [],
                    [],
                    "2026년도 수입식품통합정보시스템 고도화",
                    dump_root / "temp",
                    set(),
                    revision_metadata={
                        "revision_date": "2026-00-00",
                        "author": "송아름",
                        "approval_author": "임채현",
                    },
                )

            self.assertEqual([{"status": "updated", "project_only": False}], items)
            apply_mock.assert_called_once()
            self.assertTrue(apply_mock.call_args.kwargs["rename_files"])
            project_only_mock.assert_not_called()
        finally:
            shutil.rmtree(dump_root, ignore_errors=True)

    def test_unmatched_project_title_updates_skip_selected_candidate_renamed_path(self) -> None:
        dump_root = Path.cwd() / ".test-artifacts" / uuid4().hex
        original = dump_root / "folder" / "MFDS-ADT-A0401-01-UnitTestCase_v1.0.xlsx"
        renamed = dump_root / "folder" / "MFDS-ADT-A0401-01-UnitTestCase_v0.1.xlsx"
        original.parent.mkdir(parents=True, exist_ok=True)
        original.write_text("placeholder", encoding="utf-8")
        renamed.write_text("placeholder", encoding="utf-8")
        output = StandardOutput("MFDS-ADT-A0401-01", "UnitTestCase")
        selected = MatchCandidate(output, ScannedFile(original), 1.0, "selected")
        try:
            with (
                patch("output_file_check.folder_apply_ops.scan_folder", return_value=[ScannedFile(renamed)]),
                patch("output_file_check.folder_apply_ops.apply_batch_candidate") as apply_mock,
                patch("output_file_check.folder_apply_ops.write_updated_project_title") as project_only_mock,
            ):
                items = apply_unmatched_project_title_updates(
                    dump_root,
                    [output],
                    [],
                    [selected],
                    "Project Smoke",
                    dump_root / "temp",
                    set(),
                    revision_metadata={
                        "revision_date": "2026-00-00",
                        "author": "author",
                        "approval_author": "approval",
                    },
                )

            self.assertEqual([], items)
            apply_mock.assert_not_called()
            project_only_mock.assert_not_called()
        finally:
            shutil.rmtree(dump_root, ignore_errors=True)

    def test_unmatched_hwp_without_identity_does_not_trigger_content_read(self) -> None:
        dump_root = Path.cwd() / ".test-artifacts" / uuid4().hex
        source = dump_root / "templates" / "manual-form.hwp"
        source.parent.mkdir(parents=True, exist_ok=True)
        source.write_text("placeholder", encoding="utf-8")
        try:
            with (
                patch(
                    "output_file_check.folder_apply_ops.scan_folder",
                    return_value=[ScannedFile(source, None)],
                ),
                patch("output_file_check.folder_apply_ops.read_file_identity") as read_identity_mock,
                patch("output_file_check.folder_apply_ops.apply_batch_candidate") as apply_mock,
                patch("output_file_check.folder_apply_ops.write_updated_project_title") as project_only_mock,
            ):
                items = apply_unmatched_project_title_updates(
                    dump_root,
                    [],
                    [],
                    [],
                    "2026년도 수입식품통합정보시스템 고도화",
                    dump_root / "temp",
                    set(),
                    revision_metadata={
                        "revision_date": "2026-00-00",
                        "author": "author",
                        "approval_author": "approval",
                    },
                )

            self.assertEqual([], items)
            read_identity_mock.assert_not_called()
            apply_mock.assert_not_called()
            project_only_mock.assert_not_called()
        finally:
            shutil.rmtree(dump_root, ignore_errors=True)

    def test_project_title_for_update_uses_preview_match_before_position_guess(self) -> None:
        output = StandardOutput(
            output_id="MFDS-FP-01-SW규모산정양식",
            output_name="SW 규모(기능점수) 산정 양식",
        )
        identity = FileIdentity(
            project_title="SW 규모(기능점수) 산정 양식",
            document_title="2025년도 수입식품통합정보시스템 고도화",
            preview_text=(
                "SW 규모(기능점수) 산정 양식\n"
                "2025년도 수입식품통합정보시스템 고도화"
            ),
        )

        self.assertEqual(
            "2025년도 수입식품통합정보시스템 고도화",
            project_title_for_update(
                identity,
                "2026년도 수입식품통합정보시스템 고도화",
                output,
            ),
        )

    def test_apply_readme_splits_manual_check_categories(self) -> None:
        dump_root = Path.cwd() / "dump"
        item = {
            "status": "updated",
            "output_id": "LS",
            "output_name": "테일러링결과서",
            "new_path": str(dump_root / "테일러링결과서.hwpx"),
            "expected_filename": "LS-테일러링결과서_v0.1.hwpx",
            "file_name_changed": True,
            "project_only": False,
            "cover_changed": False,
            "cover_update_status": "partial",
            "cover_project_replace_count": 0,
            "cover_document_number_replace_count": 0,
            "cover_document_number_missing": True,
            "cover_warning_reasons": [
                "표지에서 문서번호 위치를 찾지 못해 내부 문서번호는 확인 필요",
                "사업명 변경 대상 텍스트를 찾지 못했습니다.",
            ],
            "initial_revision_status": "skipped",
            "initial_revision_cover_update_count": 0,
            "initial_revision_history_update_count": 0,
            "initial_revision_error": "",
        }

        readme = build_apply_readme(
            dump_root,
            Path("문서관리표준.xlsx"),
            "2026년도 수입식품통합정보시스템 고도화",
            [item],
            0,
        )

        self.assertIn("## 작성자/개정정보 확인 필요", readme)
        self.assertIn("## 머릿말/표지 ID 확인 필요", readme)
        self.assertIn("## 사업명 확인 필요", readme)
        self.assertNotIn("## 1차 파일명 미변경", readme)
        self.assertNotIn("## 표지 확인 필요 파일", readme)
        self.assertNotIn("## 요구사항별 자동 생성", readme)
        self.assertNotIn("프로젝트명 교체 0곳", readme)
        self.assertNotIn("상태 updated", readme)
        self.assertNotIn("표지/머릿말 0곳", readme)
        self.assertNotIn("개정이력 0곳", readme)
        self.assertIn("테일러링결과서.hwpx", readme)

    def test_apply_readme_includes_requirement_generation_summary_when_enabled(self) -> None:
        dump_root = Path.cwd() / "dump"
        requirement_result = RequirementGenerationResult(
            enabled=True,
            target_names=("Output",),
            target_count=2,
            created_items=[
                {"status": "created"},
                {"status": "created_with_warning"},
                {"status": "error"},
            ],
            skipped_items=[{"status": "skipped"}],
            error_items=[{"status": "error"}],
            readme_path=dump_root / "README_requirements.md",
            removed_items=[
                {"status": "removed"},
                {"status": "error"},
            ],
            folder_items=[
                {"status": "created"},
                {"status": "error"},
            ],
        )

        readme = build_apply_readme(
            dump_root,
            Path("standard.xlsx"),
            "Project",
            [],
            0,
            requirement_result=requirement_result,
        )

        self.assertIn("## 요구사항별 자동 생성", readme)
        self.assertIn("| 대상 산출물 | 2건 |", readme)
        self.assertIn("README_requirements.md", readme)
        self.assertIn("| 생성 오류 | 4건 |", readme)

    def test_apply_dumped_folder_passes_requirement_result_to_readme(self) -> None:
        dump_root = Path.cwd() / ".test-artifacts" / uuid4().hex
        dump_root.mkdir(parents=True, exist_ok=True)
        requirement_result = RequirementGenerationResult(True, (), 0, [], [], [])
        mapping = SimpleNamespace(
            standard_project_title="Project",
            outputs=[],
            path_templates=[],
            files=[],
            matches=[],
        )
        try:
            with (
                patch("output_file_check.folder_apply_ops.build_folder_mapping", return_value=mapping),
                patch("output_file_check.folder_apply_ops.start_allow_all_watcher", return_value=(None, None)),
                patch("output_file_check.folder_apply_ops.stop_allow_all_watcher"),
                patch("output_file_check.folder_apply_ops.run_requirement_generation_safely", return_value=requirement_result),
                patch("output_file_check.folder_apply_ops.remove_noise_files", return_value=0),
                patch("output_file_check.folder_apply_ops.write_apply_readme") as readme_mock,
                patch("output_file_check.folder_apply_ops.serialize_check_result", return_value={}),
            ):
                apply_dumped_folder(
                    Path("standard.pdf"),
                    dump_root,
                    {},
                    None,
                    dump_root / "temp",
                    "request",
                    log_prefix="folder_apply",
                    requirement_files=[],
                )

            self.assertIs(requirement_result, readme_mock.call_args.kwargs["requirement_result"])
        finally:
            shutil.rmtree(dump_root, ignore_errors=True)

    def test_cover_status_does_not_warn_when_project_title_is_verified(self) -> None:
        status = build_cover_status(
            old_document_number="LS",
            new_document_number="LS",
            old_project_title="",
            new_project_title="2026년도 수입식품통합정보시스템 고도화",
            project_title_replace_count=0,
            document_number_replace_count=0,
            project_title_verified=True,
        )

        self.assertNotIn("사업명", " / ".join(status["cover_warning_reasons"]))


if __name__ == "__main__":
    unittest.main()
