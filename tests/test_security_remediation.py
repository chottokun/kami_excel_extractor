import unittest
import re
from unittest.mock import patch, MagicMock
from pathlib import Path
import tempfile
import os
from src.kami_excel_extractor.converter import ExcelConverter
from src.kami_excel_extractor.document_generator import DocumentGenerator
from src.converter import convert_to_png

class TestSecurityRemediation(unittest.TestCase):
    def setUp(self):
        self.output_dir = Path("/tmp/test_output")
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.input_file = self.output_dir / "test.xlsx"
        self.input_file.touch()

    def tearDown(self):
        if self.input_file.exists():
            self.input_file.unlink()
        if self.output_dir.exists():
            import shutil
            shutil.rmtree(self.output_dir)

    @patch("subprocess.run")
    def test_excel_converter_uses_tempfile(self, mock_run):
        # Mock subprocess.run to return success
        mock_run.return_value = MagicMock(returncode=0, stdout="", stderr="")

        with patch("src.kami_excel_extractor.converter.Path.exists", return_value=True):
            converter = ExcelConverter(self.output_dir)
            try:
                converter.convert(self.input_file)
            except Exception:
                pass

        found_soffice = False
        for call in mock_run.call_args_list:
            args = call[0][0]
            if "soffice" in args:
                found_soffice = True
                user_inst_arg = [a for a in args if a.startswith("-env:UserInstallation=file://")]
                self.assertTrue(user_inst_arg, "UserInstallation must use file:// URI")
                path = user_inst_arg[0].split("file://")[1]
                # Pattern /tmp/lo_profile_XXXXXXX where XXXXXXX is random
                self.assertTrue(re.match(r"^/tmp/lo_profile_.+", path), f"Path should follow tempfile pattern: {path}")
                self.assertNotEqual(path, "/tmp/lo_profile_", "Should have a random suffix")

        self.assertTrue(found_soffice, "soffice should have been called")

    @patch("subprocess.run")
    def test_top_level_convert_to_png_uses_tempfile(self, mock_run):
        mock_run.return_value = MagicMock(returncode=0, stdout="", stderr="")

        with patch("src.converter.Path.exists", return_value=True):
            try:
                convert_to_png(self.input_file, self.output_dir)
            except Exception:
                pass

        found_soffice = False
        for call in mock_run.call_args_list:
            args = call[0][0]
            if "soffice" in args:
                found_soffice = True
                user_inst_arg = [a for a in args if a.startswith("-env:UserInstallation=file://")]
                self.assertTrue(user_inst_arg)
                path = user_inst_arg[0].split("file://")[1]
                self.assertTrue(re.match(r"^/tmp/lo_profile_.+", path), f"Path should follow tempfile pattern: {path}")

        self.assertTrue(found_soffice)

    @patch("subprocess.run")
    def test_document_generator_uses_tempfile(self, mock_run):
        mock_run.return_value = MagicMock(returncode=0, stdout="", stderr="")

        # Create a dummy image for _resolve_images_to_tmpdir to find
        media_dir = self.output_dir / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        (media_dir / "img.png").touch()

        generator = DocumentGenerator(self.output_dir)
        with patch("shutil.move"):
            generator.generate_pdf("# Test\n![img](img.png)", "test_pdf")

        found_soffice = False
        for call in mock_run.call_args_list:
            args = call[0][0]
            if "soffice" in args:
                found_soffice = True
                # The --outdir and the input HTML should be in a secure temp dir
                outdir_idx = args.index("--outdir") + 1
                outdir_path = args[outdir_idx]
                self.assertTrue(re.match(r"^/tmp/pdf_gen_.+", outdir_path), f"Path should follow tempfile pattern: {outdir_path}")

        self.assertTrue(found_soffice)

if __name__ == "__main__":
    unittest.main()
