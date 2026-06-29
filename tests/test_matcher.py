from pathlib import Path
import unittest

from output_file_check.folder_matching import find_contiguous_subpath, match_files_by_folder_path
from output_file_check.matcher import score_file
from output_file_check.models import FileIdentity, ScannedFile, StandardOutput


class MatcherIgnoresOutputIdTest(unittest.TestCase):
    def test_filename_output_id_does_not_match_wrong_output(self) -> None:
        output = StandardOutput("MFDS-PMC-06", "OtherDocument")
        file = ScannedFile(Path("MFDS-PMC-06-RiskRegister_v1.0.xlsx"))

        self.assertIsNone(score_file(output, file))

    def test_filename_output_name_matches_even_when_output_id_changed(self) -> None:
        output = StandardOutput("MFDS-PMC-99", "RiskRegister")
        file = ScannedFile(Path("MFDS-PMC-06-RiskRegister_v1.0.xlsx"))

        candidate = score_file(output, file)

        self.assertIsNotNone(candidate)
        self.assertEqual("파일명에 산출물명 포함", candidate.reason)

    def test_filename_parenthetical_output_name_matches_plain_standard_name(self) -> None:
        output = StandardOutput("MFDS-PMC-02", "요구사항추적표")
        file = ScannedFile(Path("MFDS-PMC-02-요구사항추적표(검사기준포함)_v1.0.xlsx"))

        candidate = score_file(output, file)

        self.assertIsNotNone(candidate)

    def test_cover_output_id_does_not_match_wrong_output(self) -> None:
        output = StandardOutput("MFDS-PMC-06", "OtherDocument")
        file = ScannedFile(
            Path("template.xlsx"),
            FileIdentity(preview_text="MFDS-PMC-06"),
        )

        self.assertIsNone(score_file(output, file))

    def test_folder_matching_prefers_name_not_old_id(self) -> None:
        wrong_by_id = StandardOutput("MFDS-PMC-06", "OtherDocument")
        right_by_name = StandardOutput("MFDS-PMC-99", "RiskRegister")
        file = ScannedFile(Path("MFDS-PMC-06-RiskRegister_v1.0.xlsx"))

        matches = match_files_by_folder_path(
            [wrong_by_id, right_by_name],
            [file],
            [],
            Path("."),
            threshold=0.72,
            folder_policy=None,
        )
        matched = {match.output.output_name: match.candidates for match in matches}

        self.assertEqual((), matched["OtherDocument"])
        self.assertEqual(file.path, matched["RiskRegister"][0].file.path)

    def test_numbered_project_end_folder_matches_standard_template(self) -> None:
        self.assertEqual(
            0,
            find_contiguous_subpath(("3.프로젝트 종료",), ("03.프로젝트 종료",)),
        )


if __name__ == "__main__":
    unittest.main()
