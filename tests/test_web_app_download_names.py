import unittest

from web_app import generated_download_name_from_template, unique_download_name


class WebAppDownloadNameTest(unittest.TestCase):
    def test_tc_download_names_follow_unit_test_template_name(self):
        self.assertEqual(
            "MFDS-ADT-A0401-01-단위시험케이스_SFR-IIL-002_v0.1.hwpx",
            generated_download_name_from_template(
                "MFDS-ADT-A0401-01-단위시험케이스_SFR-IIL-002_v0.1.hwpx",
                ".hwpx",
            ),
        )
        self.assertEqual(
            "MFDS-ADT-A0401-01-단위시험케이스_SFR-IIL-002_v0.1.xlsx",
            generated_download_name_from_template(
                "MFDS-ADT-A0401-01-단위시험케이스_SFR-IIL-002_v0.1.hwpx",
                ".xlsx",
            ),
        )

    def test_ts_download_name_follows_scenario_template_name(self):
        self.assertEqual(
            "MFDS-ADT-A0401-02-통합시험시나리오_SFR-IIL-002_v0.1.xlsx",
            generated_download_name_from_template(
                "MFDS-ADT-A0401-02-통합시험시나리오_SFR-IIL-002_v0.1.xlsx",
                ".xlsx",
            ),
        )

    def test_duplicate_template_based_names_are_disambiguated(self):
        existing_names: set[str] = set()
        self.assertEqual("단위시험케이스.hwpx", unique_download_name(existing_names, "단위시험케이스.hwpx"))
        self.assertEqual("단위시험케이스_2.hwpx", unique_download_name(existing_names, "단위시험케이스.hwpx"))


if __name__ == "__main__":
    unittest.main()
