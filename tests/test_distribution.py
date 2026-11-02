"""Keep personal files out of portable releases."""

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

from scripts import build_zip


class DistributionTests(unittest.TestCase):
    def test_zip_contains_only_project_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in build_zip.RELEASE_FILES:
                (root / name).write_text("# example", encoding="utf-8")
            (root / "__init__.py").write_text('__version__ = "3.1.0"', encoding="utf-8")
            for name in (".env", "config.local.env", "paper.docx", "private.json"):
                (root / name).write_text("must not ship", encoding="utf-8")
            with patch.object(build_zip, "ROOT", root):
                path = build_zip.build_zip(root / "dist")
            with zipfile.ZipFile(path) as archive:
                names = {Path(name).name for name in archive.namelist()}
            self.assertEqual(names, set(build_zip.RELEASE_FILES))


if __name__ == "__main__":
    unittest.main()
