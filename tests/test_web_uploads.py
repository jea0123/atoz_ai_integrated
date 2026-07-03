from pathlib import Path
import shutil
import unittest
from uuid import uuid4

from web_uploads import save_check_uploads


class WebUploadFilteringTest(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(".test-artifacts") / f"web-uploads-{uuid4().hex}"
        self.root.mkdir(parents=True, exist_ok=True)
        self.addCleanup(lambda: shutil.rmtree(self.root, ignore_errors=True))

    def test_check_uploads_keep_supported_documents_and_skip_noise(self) -> None:
        standard_path, folder_path = save_check_uploads(
            self.root,
            {
                "standard_file": [("standard.pdf", b"standard")],
                "folder_files": [
                    ("project/docs/keep.hwpx", b"keep"),
                    ("project/root.xlsx", b"keep"),
                    ("project/bak/skip.hwpx", b"skip"),
                    ("project/old/skip.docx", b"skip"),
                    ("project/readme.txt", b"skip"),
                    ("project/image.png", b"skip"),
                    ("project/manual.pdf", b"skip"),
                ],
            },
        )

        self.assertTrue(standard_path.exists())
        self.assertTrue((folder_path / "project" / "docs" / "keep.hwpx").exists())
        self.assertTrue((folder_path / "project" / "root.xlsx").exists())
        self.assertFalse((folder_path / "project" / "bak" / "skip.hwpx").exists())
        self.assertFalse((folder_path / "project" / "old" / "skip.docx").exists())
        self.assertFalse((folder_path / "project" / "readme.txt").exists())
        self.assertFalse((folder_path / "project" / "image.png").exists())
        self.assertFalse((folder_path / "project" / "manual.pdf").exists())


if __name__ == "__main__":
    unittest.main()
