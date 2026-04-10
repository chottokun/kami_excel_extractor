import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

@patch("shutil.which", side_effect=lambda x: f"/usr/bin/{x}")
@patch("kami_excel_extractor.document_generator.subprocess.run")
@patch("shutil.move")
def test_generate_pdf_path_traversal_remediation(mock_move, mock_run, mock_which, doc_gen, tmp_path):
    """Path traversal attempt via output_name is blocked by secure_filename."""
    mock_run.return_value = MagicMock(returncode=0)

    # Malicious output name
    malicious_name = "../../../tmp/evil"
    md_content = "# Security Test"

    # Mocking Path.exists to return True for the expected PDF in temp dir
    with patch("pathlib.Path.exists", return_value=True):
        result = doc_gen.generate_pdf(md_content, malicious_name)

    # The result should be under tmp_path, and its name should be sanitized
    # secure_filename("../../../tmp/evil") should result in something like "tmp_evil"
    assert result is not None
    assert result.parent == tmp_path
    assert ".." not in result.name
    assert result.name == "tmp_evil.pdf"

    # Ensure it's inside the output directory
    assert result.resolve().is_relative_to(tmp_path.resolve())

@patch("shutil.which", side_effect=lambda x: f"/usr/bin/{x}")
def test_generate_pdf_with_spaces_and_dots(mock_which, doc_gen, tmp_path):
    """Test that filenames with spaces and dots are handled correctly."""
    output_name = "my report v1.0"
    md_content = "# Title"

    with patch("kami_excel_extractor.document_generator.subprocess.run") as mock_run, \
         patch("shutil.move"), \
         patch("pathlib.Path.exists", return_value=True):
        mock_run.return_value = MagicMock(returncode=0)
        result = doc_gen.generate_pdf(md_content, output_name)

    assert result is not None
    assert result.name == "my_report_v1.0.pdf"
