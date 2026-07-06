import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from qa_generation import generate_tc
from qa_generation.generate_tc import extract_text_from_pdf, generate_test_cases


class FakePdfDocument:
    def __init__(self, text: str):
        self.text = text

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def __len__(self):
        return 1

    def load_page(self, _page_num):
        return SimpleNamespace(get_text=lambda: self.text)


class GenerateTcTextExtractionTest(unittest.TestCase):
    def build_cases_from_block(
        self,
        block_text,
        screen_id="UI-SFD-001-01-01",
        unit_test_id="UT-SFD-001-01-01",
        progress_callback=None,
        block_status_callback=None,
    ):
        return generate_tc.build_test_cases_from_text(
            extracted_text=block_text,
            model_name="model",
            ollama_url="http://localhost",
            screen_blocks=[{
                "screen_id": screen_id,
                "unit_test_id": unit_test_id,
                "text": block_text,
            }],
            progress_callback=progress_callback,
            block_status_callback=block_status_callback,
        )

    def test_hwp_design_is_not_supported_yet(self):
        with tempfile.TemporaryDirectory() as temp:
            document_path = Path(temp) / "SFR-001_사용자인터페이스설계서.hwp"
            document_path.write_bytes(b"placeholder")

            with patch("qa_generation.generate_tc.extract_document_text") as extract_text:
                fake_fitz = SimpleNamespace(open=lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("MuPDF should not open HWP")))
                with patch("qa_generation.generate_tc.fitz", fake_fitz):
                    with self.assertRaisesRegex(ValueError, "PDF"):
                        extract_text_from_pdf(document_path)

            extract_text.assert_not_called()

    def test_pdf_design_uses_mupdf_first(self):
        with tempfile.TemporaryDirectory() as temp:
            document_path = Path(temp) / "SFR-001_사용자인터페이스설계서.pdf"
            document_path.write_bytes(b"placeholder")

            with patch("qa_generation.generate_tc.extract_document_text") as extract_text:
                fake_fitz = SimpleNamespace(open=lambda *_args, **_kwargs: FakePdfDocument("UI-SFR-001-01"))
                with patch("qa_generation.generate_tc.fitz", fake_fitz):
                    result = extract_text_from_pdf(document_path)

            self.assertEqual("UI-SFR-001-01", result)
            extract_text.assert_not_called()

    def test_pdf_design_falls_back_to_document_text_when_mupdf_is_empty(self):
        with tempfile.TemporaryDirectory() as temp:
            document_path = Path(temp) / "SFR-001_사용자인터페이스설계서.pdf"
            document_path.write_bytes(b"placeholder")

            with patch("qa_generation.generate_tc.extract_document_text", return_value="UI-SFR-001-02") as extract_text:
                fake_fitz = SimpleNamespace(open=lambda *_args, **_kwargs: FakePdfDocument(""))
                with patch("qa_generation.generate_tc.fitz", fake_fitz):
                    result = extract_text_from_pdf(document_path)

            self.assertEqual("UI-SFR-001-02", result)
            extract_text.assert_called_once_with(document_path)

    def test_generate_test_cases_reports_ai_timeout_reason(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            pdf_path = root / "SFR-001_사용자인터페이스설계서.pdf"
            template_path = root / "SFR-001_단위시험케이스.hwpx"
            pdf_path.write_bytes(b"placeholder")
            template_path.write_bytes(b"placeholder")
            screen_blocks = [{
                "screen_id": "UI-SFR-001-01",
                "unit_test_id": "UT-SFR-001-01",
                "text": "화면ID UI-SFR-001-01\n처리흐름\n① 조회 버튼을 클릭한다.",
            }]

            with patch("qa_generation.generate_tc.extract_cover_author_from_document", return_value=""):
                with patch(
                    "qa_generation.generate_tc.call_ollama",
                    side_effect=generate_tc.requests.exceptions.Timeout(),
                ):
                    result = generate_test_cases(
                        pdf_path=pdf_path,
                        model_name="model",
                        ollama_url="http://localhost",
                        output_dir=root / "out",
                        template_path=template_path,
                        extracted_text=screen_blocks[0]["text"],
                        screen_blocks=screen_blocks,
                    )

            self.assertFalse(result["ok"])
            self.assertIn("AI 응답 시간 초과", result["error"])
            self.assertIn("UI-SFR-001-01", result["error"])

    def test_document_size_classification_uses_screen_id_count(self):
        self.assertEqual("small", generate_tc.classify_document_size(5))
        self.assertEqual("medium", generate_tc.classify_document_size(6))
        self.assertEqual("medium", generate_tc.classify_document_size(10))
        self.assertEqual("large", generate_tc.classify_document_size(11))

    def test_block_size_classification_uses_flow_count_and_text_length(self):
        self.assertEqual("small", generate_tc.classify_block_size(3, "x" * 1499))
        self.assertEqual("medium", generate_tc.classify_block_size(4, "x"))
        self.assertEqual("medium", generate_tc.classify_block_size(1, "x" * 1500))
        self.assertEqual("large", generate_tc.classify_block_size(9, "x"))
        self.assertEqual("large", generate_tc.classify_block_size(1, "x" * 3000))

    def test_llm_limits_follow_block_size_classification(self):
        self.assertEqual((1024, 60), generate_tc.get_llm_limits(3, "x" * 1499))
        self.assertEqual((2048, 90), generate_tc.get_llm_limits(4, "x"))
        self.assertEqual((4096, 180), generate_tc.get_llm_limits(9, "x"))

    def test_progress_message_includes_actual_llm_request_limits(self):
        block_text = """4.1. 내부사용자관리
화면ID UI-SFD-001-01-01
화면명 내부사용자관리
처리흐름
① 검색 버튼을 클릭한다.
"""
        messages = []
        ai_response = """{
          "test_cases": [
            {
              "순서": 1,
              "테스트_케이스": "검색 버튼 클릭",
              "예상_결과": "목록이 조회된다."
            }
          ]
        }"""

        with patch("qa_generation.generate_tc.call_ollama", return_value=ai_response):
            self.build_cases_from_block(block_text, progress_callback=messages.append)

        joined_messages = "\n".join(messages)
        self.assertIn("num_predict=1024", joined_messages)
        self.assertIn("timeout=60s", joined_messages)

    def test_block_status_callback_reports_request_limits_and_statuses(self):
        block_text = """4.1. 내부사용자관리
화면ID UI-SFD-001-01-01
화면명 내부사용자관리
처리흐름
① 검색 버튼을 클릭한다.
"""
        events = []
        ai_response = """{
          "test_cases": [
            {
              "순서": 1,
              "테스트_케이스": "검색 버튼 클릭",
              "예상_결과": "목록이 조회된다."
            }
          ]
        }"""

        with patch("qa_generation.generate_tc.call_ollama", return_value=ai_response):
            self.build_cases_from_block(block_text, block_status_callback=events.append)

        self.assertEqual(["queued", "running", "updated"], [event["status"] for event in events])
        self.assertEqual("UI-SFD-001-01-01", events[-1]["screen_id"])
        self.assertEqual("small", events[-1]["block_size"])
        self.assertEqual(1024, events[-1]["num_predict"])
        self.assertEqual(60, events[-1]["timeout"])
        self.assertEqual(1, events[-1]["generated_count"])

    def test_small_blocks_are_processed_first_but_results_keep_original_order(self):
        large_block = """4.1. 큰 화면
화면ID UI-LARGE-001
화면명 큰 화면
처리흐름
① 첫 번째 작업을 수행한다.
② 두 번째 작업을 수행한다.
③ 세 번째 작업을 수행한다.
④ 네 번째 작업을 수행한다.
"""
        small_block = """4.2. 작은 화면
화면ID UI-SMALL-001
화면명 작은 화면
처리흐름
① 검색 버튼을 클릭한다.
"""
        screen_blocks = [
            {
                "screen_id": "UI-LARGE-001",
                "unit_test_id": "UT-LARGE-001",
                "text": large_block,
            },
            {
                "screen_id": "UI-SMALL-001",
                "unit_test_id": "UT-SMALL-001",
                "text": small_block,
            },
        ]
        call_order = []

        def fake_call(_url, _model, _system_prompt, user_prompt, **_kwargs):
            if "UI-SMALL-001" in user_prompt:
                call_order.append("UI-SMALL-001")
                return '{"test_cases":[{"순서":1,"테스트_케이스":"작은 화면 처리","예상_결과":"성공"}]}'
            call_order.append("UI-LARGE-001")
            return '{"test_cases":[{"순서":1,"테스트_케이스":"큰 화면 처리","예상_결과":"성공"}]}'

        with patch("qa_generation.generate_tc.call_ollama", side_effect=fake_call):
            cases = generate_tc.build_test_cases_from_text(
                extracted_text=f"{large_block}\n{small_block}",
                model_name="model",
                ollama_url="http://localhost",
                screen_blocks=screen_blocks,
            )

        self.assertEqual(["UI-SMALL-001", "UI-LARGE-001"], call_order)
        self.assertEqual(["UT-LARGE-001", "UT-SMALL-001"], [case["단위시험_ID"] for case in cases])
        self.assertNotIn("__block_original_index", cases[0])

    def test_unit_test_name_uses_screen_name_with_improvement_when_title_matches_screen(self):
        block_text = """4.1. 내부사용자관리
요구사항ID SFR-SFD-001-01
화면ID UI-SFD-001-01-01
화면명 내부사용자관리
화면설명 내부사용자를 관리하는 등록, 수정하는 화면
메뉴경로 시스템관리 > 사용자관리 > 내부사용자관리
개발구분 개선
처리흐름
① 검색조건을 입력 후 검색 버튼을 클릭한다.
② 내부사용자 목록이 검색된다.
"""
        ai_response = """{
          "test_cases": [
            {
              "화면명": "UT-SFD-001-01-01",
              "단위시험_명": "UT-SFD-001-01-01",
              "사전조건": "로그인 상태",
              "순서": 1,
              "테스트_케이스": "검색 버튼 클릭",
              "테스트_데이터": "AI가 만든 값",
              "예상_결과": "목록이 조회된다."
            }
          ]
        }"""

        with patch("qa_generation.generate_tc.call_ollama", return_value=ai_response):
            cases = self.build_cases_from_block(block_text)

        self.assertEqual("내부사용자관리", cases[0]["단위시험_제목"])
        self.assertEqual("내부사용자관리 개선", cases[0]["단위시험_명"])
        self.assertEqual("내부사용자관리", cases[0]["화면명"])
        self.assertEqual("", cases[0]["테스트_데이터"])

    def test_test_case_text_strips_leading_sequence_marker(self):
        block_text = """4.1. 내부사용자관리
요구사항ID SFR-SFD-001-01
화면ID UI-SFD-001-01-01
화면명 내부사용자관리
처리흐름
① 검색조건을 입력 후 검색 버튼을 클릭한다.
② 내부사용자 목록 중 하나의 행을 클릭한다.
"""
        ai_response = """{
          "test_cases": [
            {
              "순서": 1,
              "테스트_케이스": "1. 검색조건을 입력 후 검색 버튼을 클릭한다.",
              "예상_결과": "목록이 조회된다."
            },
            {
              "순서": 2,
              "테스트_케이스": "② 내부사용자 목록 중 하나의 행을 클릭한다.",
              "예상_결과": "상세 정보가 조회된다."
            }
          ]
        }"""

        with patch("qa_generation.generate_tc.call_ollama", return_value=ai_response):
            cases = self.build_cases_from_block(block_text)

        self.assertEqual("검색조건을 입력 후 검색 버튼을 클릭한다.", cases[0]["테스트_케이스"])
        self.assertEqual("내부사용자 목록 중 하나의 행을 클릭한다.", cases[1]["테스트_케이스"])

    def test_unit_test_name_uses_section_title_when_title_is_more_specific_than_screen(self):
        block_text = """4.1. 전자위생증명서 연계 국가 확대
요구사항ID SFR-IIL-002-01
화면ID UI-IIL-002-01-01-01
화면명 전자위생증명서(수입업무)
화면설명 내부담당자가 수입전자위생증명서 목록을 확인할 수 있는 화면
메뉴경로 시스템관리 > 연계관리 > 전자위생증명서(수입업무)
개발구분 개선
처리흐름
① 수입증명서 대장 리스트 화면에서 증명서를 확인하고 싶은 행을 클릭한다.
② 보고서 버튼을 클릭한다.
"""
        ai_response = """{
          "test_cases": [
            {
              "화면명": "UT-IIL-002-01-01-01",
              "단위시험_명": "UT-IIL-002-01-01-01",
              "사전조건": "로그인 상태",
              "순서": 1,
              "테스트_케이스": "행 선택",
              "예상_결과": "상세 정보가 표시된다."
            }
          ]
        }"""

        with patch("qa_generation.generate_tc.call_ollama", return_value=ai_response):
            cases = self.build_cases_from_block(
                block_text,
                screen_id="UI-IIL-002-01-01-01",
                unit_test_id="UT-IIL-002-01-01-01",
            )

        self.assertEqual("전자위생증명서 연계 국가 확대", cases[0]["단위시험_제목"])
        self.assertEqual("전자위생증명서 연계 국가 확대", cases[0]["단위시험_명"])
        self.assertEqual("전자위생증명서(수입업무)", cases[0]["화면명"])


if __name__ == "__main__":
    unittest.main()
