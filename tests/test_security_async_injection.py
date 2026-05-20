import asyncio
import subprocess
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from kami_excel_extractor.document_generator import DocumentGenerator


@pytest.mark.asyncio
async def test_agenerate_pdf_uses_absolute_paths(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    generator = DocumentGenerator(output_dir)

    # Mock _get_soffice_path to return a relative path to see if it gets resolved at call site
    with (
        patch.object(DocumentGenerator, "_get_soffice_path", return_value="soffice"),
        patch("asyncio.create_subprocess_exec", new_callable=AsyncMock) as mock_exec,
        patch("pathlib.Path.exists", return_value=True),
        patch("pathlib.Path.rglob", return_value=[tmp_path / "test.pdf"]),
        patch("shutil.move"),
    ):
        # Mock communicate and wait
        mock_proc = AsyncMock()
        mock_proc.communicate.return_value = (b"", b"")
        mock_proc.wait.return_value = 0
        mock_proc.returncode = 0
        mock_exec.return_value = mock_proc

        await generator.agenerate_pdf("# Test", "test_report")

        assert mock_exec.called
        args = mock_exec.call_args[0]

        # Check executable
        assert Path(args[0]).is_absolute(), f"Executable path {args[0]} should be absolute"

        # Check other arguments that are paths
        outdir = args[5]
        html_path = args[6]

        assert Path(outdir).is_absolute(), f"Outdir {outdir} should be absolute"
        assert Path(html_path).is_absolute(), f"HTML path {html_path} should be absolute"


@pytest.mark.asyncio
async def test_arun_soffice_conversion_absolute_paths():
    # Targeted test for _arun_soffice_conversion
    with patch("asyncio.create_subprocess_exec", new_callable=AsyncMock) as mock_exec:
        mock_proc = AsyncMock()
        mock_proc.communicate.return_value = (b"", b"")
        mock_proc.returncode = 0
        mock_exec.return_value = mock_proc

        output_dir = Path("test_out").resolve()
        generator = DocumentGenerator(output_dir)

        tmp_dir = Path("relative_tmp")
        temp_html = Path("relative.html")

        # We need to mock _get_soffice_path to return something relative
        with patch.object(DocumentGenerator, "_get_soffice_path", return_value="soffice"):
            # We also need to mock Path.exists for the expected pdf
            with patch("pathlib.Path.exists", return_value=True):
                await generator._arun_soffice_conversion(tmp_dir, temp_html)

                assert mock_exec.called
                args = mock_exec.call_args[0]

                # soffice_path
                assert Path(args[0]).is_absolute(), f"Arg 0 {args[0]} should be absolute"
                # --outdir
                assert Path(args[5]).is_absolute(), f"Arg 5 {args[5]} should be absolute"
                # input html
                assert Path(args[6]).is_absolute(), f"Arg 6 {args[6]} should be absolute"


def test_run_soffice_conversion_absolute_paths():
    # Targeted test for _run_soffice_conversion
    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 0

        output_dir = Path("test_out").resolve()
        generator = DocumentGenerator(output_dir)

        tmp_dir = Path("relative_tmp")
        temp_html = Path("relative.html")

        with patch.object(DocumentGenerator, "_get_soffice_path", return_value="soffice"):
            with patch("pathlib.Path.exists", return_value=True):
                generator._run_soffice_conversion(tmp_dir, temp_html)

                assert mock_run.called
                args = mock_run.call_args[0][0]

                assert Path(args[0]).is_absolute(), f"Arg 0 {args[0]} should be absolute"
                assert Path(args[5]).is_absolute(), f"Arg 5 {args[5]} should be absolute"
                assert Path(args[6]).is_absolute(), f"Arg 6 {args[6]} should be absolute"
