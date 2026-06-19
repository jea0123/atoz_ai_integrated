from pathlib import Path
import unittest
from unittest.mock import patch

from output_file_check.models import StandardOutput
from output_file_check.requirement_generation import (
    build_generated_filename,
    extract_requirement_ids,
    generate_requirement_documents,
    select_applied_template_paths,
    should_require_template_requirement_id,
    template_has_requirement_id_tail,
)


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

    def test_extract_requirement_ids_ignores_non_sfr_ids(self) -> None:
        self.assertEqual((), extract_requirement_ids("REQ-001.txt"))
        self.assertEqual((), extract_requirement_ids("MFDS-PMC-06-RiskRegister_UIR-001_v1.0.xlsx"))
        self.assertEqual(("SFR-001",), extract_requirement_ids("MFDS-PMC-06-RiskRegister_SFR-001_v1.0.xlsx"))

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


if __name__ == "__main__":
    unittest.main()
