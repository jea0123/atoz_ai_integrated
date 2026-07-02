import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from qa_generation.folder_pipeline import (
    build_requirement_work_items,
    create_uploaded_qa_source_dump,
    preview_folder_qa_matching,
    select_qa_source_files,
)
from web_app import resolve_qa_folder_target


def make_file(root: Path, name: str) -> Path:
    path = root / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("", encoding="utf-8")
    return path


class QaFolderSelectionTest(unittest.TestCase):
    def test_full_alternative_source_folder_does_not_use_dump_root_candidates(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"
            alt_root = root / "alternative"

            make_file(dump_root, "SFR-001_사용자인터페이스설계서.pdf")
            make_file(dump_root, "SFR-001_단위시험케이스.hwpx")
            make_file(dump_root, "SFR-001_단위시험결과서.hwpx")
            make_file(dump_root, "SFR-001_통합시험시나리오.xlsx")
            make_file(dump_root, "SFR-001_통합시험결과서.xlsx")

            alt_files = [
                make_file(alt_root, "SFR-001_사용자인터페이스설계서.pdf"),
                make_file(alt_root, "SFR-001_단위시험케이스.hwpx"),
                make_file(alt_root, "SFR-001_단위시험결과서.hwpx"),
                make_file(alt_root, "SFR-001_통합시험시나리오.xlsx"),
                make_file(alt_root, "SFR-001_통합시험결과서.xlsx"),
            ]

            selection = select_qa_source_files(
                dump_root,
                [],
                qa_source_paths=alt_files,
                qa_source_is_override=True,
            )

            for role in (
                "ui_design",
                "tc_template",
                "unit_result_template",
                "ts_template",
                "integration_result_template",
            ):
                selected = Path(selection[role]["by_requirement"]["SFR-001"]["path"])
                self.assertIn(alt_root, selected.parents)

    def test_role_without_alternative_source_still_falls_back_to_dump_root(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"
            alt_root = root / "alternative"

            make_file(dump_root, "SFR-001_단위시험케이스.hwpx")
            ui_design = make_file(alt_root, "SFR-001_사용자인터페이스설계서.pdf")

            selection = select_qa_source_files(
                dump_root,
                [ui_design],
            )

            selected_ui = Path(selection["ui_design"]["by_requirement"]["SFR-001"]["path"])
            selected_tc = Path(selection["tc_template"]["by_requirement"]["SFR-001"]["path"])

            self.assertIn(alt_root, selected_ui.parents)
            self.assertIn(dump_root, selected_tc.parents)

    def test_original_named_folder_is_included_in_candidate_scan(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"
            make_file(dump_root / "원본", "SFR-001_통합시험시나리오.xlsx")

            selection = select_qa_source_files(dump_root, [])
            selected_ts = Path(selection["ts_template"]["by_requirement"]["SFR-001"]["path"])

            self.assertIn("원본", selected_ts.parts)

    def test_ui_design_matching_uses_pdf_only(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"

            make_file(dump_root, "SFR-001_사용자인터페이스설계서.hwp")
            make_file(dump_root, "SFR-002_사용자인터페이스설계서.hwpx")
            pdf_design = make_file(dump_root, "SFR-003_사용자인터페이스설계서.pdf")

            selection = select_qa_source_files(dump_root, [])
            ui_by_requirement = selection["ui_design"]["by_requirement"]

            self.assertNotIn("SFR-001", ui_by_requirement)
            self.assertNotIn("SFR-002", ui_by_requirement)
            self.assertEqual(str(pdf_design), ui_by_requirement["SFR-003"]["path"])

    def test_uploaded_source_files_do_not_replace_dump_root_artifact_targets(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"
            uploaded_root = root / "uploaded"

            make_file(dump_root, "SFR-001_단위시험케이스.hwpx")
            uploaded_tc = make_file(uploaded_root, "SFR-001_단위시험케이스.hwpx")

            selection = select_qa_source_files(
                dump_root,
                [],
                qa_source_paths=[uploaded_tc],
                qa_source_is_override=False,
            )

            selected_tc = Path(selection["tc_template"]["by_requirement"]["SFR-001"]["path"])

            self.assertIn(dump_root, selected_tc.parents)

    def test_document_specific_folder_is_input_only_for_artifact_placement(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"
            tc_input_root = root / "tc-input"

            make_file(dump_root, "SFR-001_사용자인터페이스설계서.pdf")
            target_tc = make_file(dump_root, "SFR-001_단위시험케이스.hwpx")
            make_file(dump_root, "SFR-001_단위시험결과서.hwpx")
            make_file(dump_root, "SFR-001_통합시험시나리오.xlsx")
            make_file(dump_root, "SFR-001_통합시험결과서.xlsx")
            input_tc = make_file(tc_input_root, "SFR-001_단위시험케이스.hwpx")

            selection = select_qa_source_files(
                dump_root,
                [],
                tc_source_paths=[input_tc],
            )
            selected_tc = selection["tc_template"]["by_requirement"]["SFR-001"]
            work_items = build_requirement_work_items(selection)

            self.assertEqual(str(input_tc), selected_tc["path"])
            self.assertEqual(str(target_tc), selected_tc["placement_path"])
            self.assertEqual(str(input_tc), work_items[0]["tc_template_path"])
            self.assertEqual(str(target_tc), work_items[0]["tc_template_target_path"])

    def test_document_specific_artifact_input_requires_target_in_selected_folder(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"
            tc_input_root = root / "tc-input"

            make_file(dump_root, "SFR-001_사용자인터페이스설계서.pdf")
            input_tc = make_file(tc_input_root, "SFR-001_단위시험케이스.hwpx")

            selection = select_qa_source_files(
                dump_root,
                [],
                tc_source_paths=[input_tc],
            )

            self.assertNotIn("SFR-001", selection["tc_template"]["by_requirement"])

    def test_preview_folder_qa_matching_returns_matched_files_without_generation(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            dump_root = root / "dump"

            make_file(dump_root, "SFR-001_사용자인터페이스설계서.pdf")
            make_file(dump_root, "SFR-001_단위시험케이스.hwpx")
            make_file(dump_root, "SFR-001_단위시험결과서.hwpx")
            make_file(dump_root, "SFR-001_통합시험시나리오.xlsx")
            make_file(dump_root, "SFR-001_통합시험결과서.xlsx")

            payload = preview_folder_qa_matching(dump_root)

            self.assertTrue(payload["ok"])
            self.assertTrue(payload["match_preview"])
            self.assertEqual(1, payload["requirement_count"])
            self.assertEqual(5, len(payload["source_files"]))
            self.assertEqual([], payload["placed_files"])

    def test_uploaded_source_folder_is_materialized_as_versioned_result_folder(self):
        with tempfile.TemporaryDirectory() as temp:
            dump_parent = Path(temp) / "qa-folder-dumps"
            dump_root = create_uploaded_qa_source_dump(
                [
                    ("원본/SFR-001_단위시험케이스.hwpx", b"tc"),
                    ("원본/하위/SFR-001_통합시험시나리오.xlsx", b"ts"),
                ],
                dump_parent=dump_parent,
            )

            self.assertIsNotNone(dump_root)
            assert dump_root is not None
            self.assertEqual("원본_v0.1", dump_root.name)
            self.assertTrue((dump_root / "SFR-001_단위시험케이스.hwpx").exists())
            self.assertTrue((dump_root / "하위" / "SFR-001_통합시험시나리오.xlsx").exists())

    def test_qa_target_prefers_alternative_path_over_dump_and_upload(self):
        with patch("web_app.create_uploaded_qa_source_dump") as create_dump:
            effective, qa_source_root, upload_items = resolve_qa_folder_target(
                "C:/dump",
                "C:/alternative",
                [("원본/SFR-001_단위시험케이스.hwpx", b"tc")],
            )

        create_dump.assert_not_called()
        self.assertEqual("C:/alternative", effective)
        self.assertEqual("C:/alternative", qa_source_root)
        self.assertEqual([], upload_items)

    def test_qa_target_prefers_upload_over_dump(self):
        uploaded_dump = Path("C:/qa-folder-dumps/원본_v0.1")
        with patch("web_app.create_uploaded_qa_source_dump", return_value=uploaded_dump):
            effective, qa_source_root, upload_items = resolve_qa_folder_target(
                "C:/dump",
                "",
                [("원본/SFR-001_단위시험케이스.hwpx", b"tc")],
            )

        self.assertEqual(str(uploaded_dump), effective)
        self.assertEqual(str(uploaded_dump), qa_source_root)
        self.assertEqual([], upload_items)

    def test_qa_target_uses_dump_when_no_override_exists(self):
        effective, qa_source_root, upload_items = resolve_qa_folder_target("C:/dump", "", [])

        self.assertEqual("C:/dump", effective)
        self.assertEqual("", qa_source_root)
        self.assertEqual([], upload_items)


if __name__ == "__main__":
    unittest.main()
