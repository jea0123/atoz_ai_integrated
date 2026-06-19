from pathlib import Path
import unittest

from output_file_check.standard_reader import extract_output_section, read_standard_outputs


class StandardReaderManagementSectionTest(unittest.TestCase):
    def test_management_section_marker_does_not_require_star_or_space(self) -> None:
        text = "\n".join(
            [
                "문서관리표준",
                "관리문서ID",
                "MFDS-PMC-06-위험관리대장",
                "3.1.2 파일명",
                "MFDS-ADT-A01-개발문서",
            ]
        )

        section = extract_output_section(text, "management")

        self.assertIn("MFDS-PMC-06-위험관리대장", section)
        self.assertNotIn("MFDS-ADT-A01-개발문서", section)

    def test_management_outputs_read_output_name_column_and_arbitrary_id_column(self) -> None:
        text = "\n".join(
            [
                "문서관리표준",
                "* 관리문서 ID",
                "구분",
                "프로세스 코드",
                "산출물명",
                "산출물ID",
                "폴더명",
                "프로젝트",
                "표준",
                "PP",
                "사업수행계획서",
                "112233",
                "프로젝트 표준",
                "문서관리표준",
                "22",
                "프로젝트 표준",
                "요구사항추적표",
                "(검사기준포함)",
                "15",
                "요구관리",
                "회의록",
                "MFDS-업무ID-YYMMDD-회의명",
                "의사소통관리",
                "문서관리표준",
                "2025년도 수입식품통합정보시스템 고도화",
                "MFDS-PP-02",
                "V1.0",
                "2025.03.26",
                "임채현",
                "문서관리표준",
                "5",
                "㈜에이투지시스템",
                "* 산출물 코드",
                "구분",
                "프로세스",
                "코드",
                "산출물명",
                "산출물ID",
                "폴더명",
                "위험관리대장",
                "20",
                "위험관리",
                "활동명",
                "작업명",
                "산출물",
                "산출물ID",
                "요구사항 분석",
                "인터뷰계획서",
                "MFDS-ADT-R0101-01-인터뷰계획서",
                "3.1.2 파일명",
            ]
        )

        outputs = read_standard_outputs(Path("standard.pdf"), text, category="management")

        self.assertEqual(
            ["사업수행계획서", "문서관리표준", "요구사항추적표(검사기준포함)", "회의록", "위험관리대장"],
            [output.output_name for output in outputs],
        )
        self.assertEqual(
            ["112233", "22", "15", "MFDS-업무ID-YYMMDD-회의명", "20"],
            [output.output_id for output in outputs],
        )

    def test_management_outputs_do_not_scan_whole_document_without_marker(self) -> None:
        text = "\n".join(
            [
                "문서관리표준",
                "A-001          사업수행계획서",
                "RISK-2026      위험관리대장",
            ]
        )

        outputs = read_standard_outputs(Path("standard.pdf"), text, category="management")

        self.assertEqual([], outputs)

    def test_development_outputs_read_output_name_column_with_arbitrary_ids(self) -> None:
        text = "\n".join(
            [
                "문서관리표준",
                "* 산출물 코드",
                "활동명",
                "작업명",
                "산출물",
                "산출물ID",
                "요구사항 분석",
                "요구사항 수집",
                "인터뷰계획서",
                "DEV-CUSTOM-100",
                "요구사항명세서",
                "MFDS-ADT-R0201-01-요구사항명세서",
                "문서관리표준",
                "2025년도 수입식품통합정보시스템 고도화",
                "MFDS-PP-02",
                "V1.0",
                "2025.03.26",
                "임채현",
                "문서관리표준",
                "6",
                "㈜에이투지시스템",
                "단위시험케이스",
                "MFDS-ADT-A0401-01-단위시험케이스",
                "3.1.2 파일명",
            ]
        )

        outputs = read_standard_outputs(Path("standard.pdf"), text, category="development")

        self.assertEqual(
            ["인터뷰계획서", "요구사항명세서", "단위시험케이스"],
            [output.output_name for output in outputs],
        )
        self.assertEqual(
            ["DEV-CUSTOM-100", "MFDS-ADT-R0201-01", "MFDS-ADT-A0401-01"],
            [output.output_id for output in outputs],
        )

    def test_development_output_name_still_splits_from_standard_id(self) -> None:
        text = "\n".join(
            [
                "문서관리표준",
                "* 산출물 코드",
                "MFDS-ADT-A0401-01-단위시험케이스",
                "3.1.2 파일명",
            ]
        )

        outputs = read_standard_outputs(Path("standard.pdf"), text, category="development")

        self.assertEqual(["MFDS-ADT-A0401-01"], [output.output_id for output in outputs])
        self.assertEqual(["단위시험케이스"], [output.output_name for output in outputs])


if __name__ == "__main__":
    unittest.main()
