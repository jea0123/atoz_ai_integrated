from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


class TemplatePageInputSeparationTest(unittest.TestCase):
    def test_management_template_page_only_has_proposal_input(self) -> None:
        html = (ROOT / "web" / "management-template.html").read_text(encoding="utf-8")

        self.assertIn('name="artifact_category" value="management"', html)
        self.assertIn('name="proposal_files"', html)
        self.assertNotIn('name="requirement_files"', html)

    def test_development_template_page_only_has_requirement_input(self) -> None:
        html = (ROOT / "web" / "development-template.html").read_text(encoding="utf-8")

        self.assertIn('name="artifact_category" value="development"', html)
        self.assertIn('name="requirement_files"', html)
        self.assertNotIn('name="proposal_files"', html)


if __name__ == "__main__":
    unittest.main()
