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


if __name__ == "__main__":
    unittest.main()
