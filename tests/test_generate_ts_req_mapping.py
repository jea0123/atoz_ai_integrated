import unittest
from unittest.mock import patch

from qa_generation.generate_ts import (
    derive_requirement_id_from_screen_id,
    extract_req_mapping_from_pdf,
    fill_req_mapping_from_screen_ids,
)


class GenerateTsReqMappingTest(unittest.TestCase):
    def test_requirement_id_can_be_derived_from_screen_id(self):
        self.assertEqual(
            "SFR-IIL-002",
            derive_requirement_id_from_screen_id("UI-IIL-002-01-01-01"),
        )

    def test_extract_req_mapping_falls_back_to_screen_id_shape(self):
        with patch(
            "qa_generation.generate_ts.extract_text_from_pdf",
            return_value="화면 ID\nUI-IIL-002-01-01-01\n화면명\n테스트",
        ):
            mapping = extract_req_mapping_from_pdf("dummy.pdf")

        self.assertEqual("SFR-IIL-002", mapping["UI-IIL-002-01-01-01"])

    def test_fill_req_mapping_from_screen_ids_adds_missing_values(self):
        mapping = fill_req_mapping_from_screen_ids(
            {},
            [{"화면_ID": "UI-IIL-002-01-01-01"}],
        )

        self.assertEqual("SFR-IIL-002", mapping["UI-IIL-002-01-01-01"])


if __name__ == "__main__":
    unittest.main()
