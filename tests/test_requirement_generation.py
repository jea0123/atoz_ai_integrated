from pathlib import Path
import shutil
import unittest
from unittest.mock import patch
from uuid import uuid4

from output_file_check.models import StandardOutput
from output_file_check.requirement_generation import (
    build_generated_filename,
    extract_requirement_ids,
    generate_requirement_documents,
    reset_program_source_root,
    select_applied_template_paths,
    should_create_requirement_source_folders,
    should_require_template_requirement_id,
    template_has_requirement_id_tail,
)
from web_uploads import extract_requirement_id_match_details_from_text, extract_requirement_ids_from_text


class RequirementGenerationTemplateSelectionTest(unittest.TestCase):
    def setUp(self) -> None:
        self.output = StandardOutput("MFDS-PMC-06", "RiskRegister")

    def test_applied_scope_requires_template_requirement_id_by_default(self) -> None:
        self.assertTrue(should_require_template_requirement_id({}, "applied"))
        self.assertFalse(should_require_template_requirement_id({}, "default"))

    def test_template_without_requirement_id_tail_is_not_eligible(self) -> None:
        path = Path("MFDS-PMC-06-RiskRegister_v1.0.xlsx")

        self.assertFalse(template_has_requirement_id_tail(self.output, path))

    def test_template_with_requirement_id_tail_is_eligible(self) -> None:
        path = Path("MFDS-PMC-06-RiskRegister_SFR-001_v1.0.xlsx")

        self.assertTrue(template_has_requirement_id_tail(self.output, path))

    def test_template_with_non_sfr_id_tail_is_not_eligible(self) -> None:
        for filename in (
            "MFDS-PMC-06-RiskRegister_REQ-001_v1.0.xlsx",
            "MFDS-PMC-06-RiskRegister_PMC-001_v1.0.xlsx",
            "MFDS-PMC-06-RiskRegister_UIR-001_v1.0.xlsx",
        ):
            with self.subTest(filename=filename):
                self.assertFalse(template_has_requirement_id_tail(self.output, Path(filename)))

    def test_management_scope_does_not_create_program_source_folders(self) -> None:
        self.assertFalse(
            should_create_requirement_source_folders(
                {
                    "artifact_category": "management",
                    "requirement_generation_create_source_folders": "true",
                }
            )
        )

    def test_development_scope_creates_program_source_folders_by_default(self) -> None:
        self.assertTrue(should_create_requirement_source_folders({"artifact_category": "development"}))

    def test_extract_requirement_ids_ignores_non_sfr_ids(self) -> None:
        self.assertEqual((), extract_requirement_ids("REQ-001.txt"))
        self.assertEqual((), extract_requirement_ids("MFDS-PMC-06-RiskRegister_UIR-001_v1.0.xlsx"))
        self.assertEqual(("SFR-001",), extract_requirement_ids("MFDS-PMC-06-RiskRegister_SFR-001_v1.0.xlsx"))

    def test_proposal_requirement_ids_do_not_join_whitespace_cells(self) -> None:
        text = "요구사항목록표 SFR OOO 13 SFR-001 SFR-ESS-002 S F R - IIL - 003"

        self.assertEqual(
            ("SFR-001", "SFR-ESS-002", "SFR-IIL-003"),
            extract_requirement_ids_from_text(text),
        )
        self.assertEqual(("SFR-OOO-13",), extract_requirement_ids_from_text("SFR-OOO-13"))

    def test_proposal_requirement_ids_ignore_shape_conflict_candidates(self) -> None:
        details = extract_requirement_id_match_details_from_text("SFR-001 SFR-002 SFR-013 SFR-OOO-13")

        self.assertEqual(
            ["SFR-001", "SFR-002", "SFR-013"],
            [item["requirement_id"] for item in details["kept"]],
        )
        self.assertEqual(["SFR-OOO-13"], [item["requirement_id"] for item in details["ignored"]])
        self.assertIn("SFR-013", details["ignored"][0]["ignore_reason"])

    def test_strict_applied_selection_skips_non_requirement_template(self) -> None:
        template = Path("MFDS-PMC-06-RiskRegister_v1.0.xlsx")
        apply_items = [
            {
                "status": "updated",
                "output_id": self.output.output_id,
                "output_name": self.output.output_name,
                "new_path": str(template),
            }
        ]

        with patch.object(Path, "exists", return_value=True):
            selected = select_applied_template_paths(
                self.output,
                apply_items,
                require_template_requirement_id=True,
            )

        self.assertEqual([], selected)

    def test_strict_applied_selection_keeps_requirement_template(self) -> None:
        template = Path("MFDS-PMC-06-RiskRegister_SFR-001_v1.0.xlsx")
        apply_items = [
            {
                "status": "updated",
                "output_id": self.output.output_id,
                "output_name": self.output.output_name,
                "new_path": str(template),
            }
        ]

        with patch.object(Path, "exists", return_value=True):
            selected = select_applied_template_paths(
                self.output,
                apply_items,
                require_template_requirement_id=True,
            )
        generated_name = build_generated_filename(
            self.output,
            "SFR-002",
            "v0.1",
            template.suffix,
            template_path=template,
        )

        self.assertEqual([template], selected)
        self.assertEqual("MFDS-PMC-06-RiskRegister_SFR-002_v0.1.xlsx", generated_name)

    def test_applied_generation_without_requirement_template_is_not_error(self) -> None:
        template = Path("MFDS-PMC-06-RiskRegister_v1.0.xlsx")
        apply_items = [
            {
                "status": "updated",
                "output_id": self.output.output_id,
                "output_name": self.output.output_name,
                "new_path": str(template),
            }
        ]

        with (
            patch.object(Path, "exists", return_value=True),
            patch(
                "output_file_check.requirement_generation.write_requirement_generation_readme",
                return_value=Path("README.md"),
            ),
        ):
            result = generate_requirement_documents(
                Path("dump"),
                [self.output],
                [],
                [Path("SFR-001.txt")],
                "",
                Path("."),
                {
                    "requirement_generation_targets": "__applied__",
                    "requirement_generation_create_source_folders": "false",
                },
                apply_items=apply_items,
            )

        self.assertEqual([], result.created_items)
        self.assertEqual([], result.error_items)

    def test_program_source_root_preserves_existing_txt_files(self) -> None:
        dump_root = Path.cwd() / ".test-artifacts" / uuid4().hex
        source_root = dump_root / "02.수행" / "프로그램소스"
        note_path = source_root / "메모.txt"
        nested_note_path = source_root / "old" / "기존.txt"
        source_root.mkdir(parents=True, exist_ok=True)
        note_path.write_text("keep", encoding="utf-8")
        nested_note_path.parent.mkdir(parents=True, exist_ok=True)
        nested_note_path.write_text("keep", encoding="utf-8")
        try:
            result = reset_program_source_root(dump_root, source_root)

            self.assertEqual([], result)
            self.assertTrue(note_path.exists())
            self.assertTrue(nested_note_path.exists())
        finally:
            shutil.rmtree(dump_root, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
