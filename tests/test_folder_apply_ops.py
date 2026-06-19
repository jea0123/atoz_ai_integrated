from pathlib import Path
import unittest

from output_file_check.folder_apply_ops import build_target_filename, validate_output_id
from output_file_check.models import StandardOutput


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


if __name__ == "__main__":
    unittest.main()
