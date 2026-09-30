"""Regression tests for collision handling in PDF renaming."""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import rename_pdfs


class RenameCollisionTests(unittest.TestCase):
    META = {"title": "Paper", "author": "Jane Smith", "year": "2024"}

    def test_already_correct_name_is_not_suffixed_or_renamed(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "Smith, Jane (2024) - Paper.pdf"
            source.write_bytes(b"placeholder")

            with patch.object(rename_pdfs, "get_metadata", return_value=self.META):
                renamed = rename_pdfs.rename_single_pdf(str(source))

            self.assertFalse(renamed)
            self.assertTrue(source.exists())
            self.assertFalse(
                (Path(directory) / "Smith, Jane (2024) - Paper (1).pdf").exists()
            )

    def test_genuine_collision_gets_suffix(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "input.pdf"
            target = Path(directory) / "Smith, Jane (2024) - Paper.pdf"
            suffixed_target = Path(directory) / "Smith, Jane (2024) - Paper (1).pdf"
            source.write_bytes(b"placeholder")
            target.write_bytes(b"existing")

            with patch.object(rename_pdfs, "get_metadata", return_value=self.META):
                renamed = rename_pdfs.rename_single_pdf(str(source))

            self.assertTrue(renamed)
            self.assertFalse(source.exists())
            self.assertEqual(target.read_bytes(), b"existing")
            self.assertEqual(suffixed_target.read_bytes(), b"placeholder")


if __name__ == "__main__":
    unittest.main()
