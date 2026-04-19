import pytest
import os
from pathlib import Path
from unittest.mock import AsyncMock, patch, MagicMock
from kami_excel_extractor.cli import run_async

@pytest.mark.asyncio
async def test_cli_output_path_traversal_remediation(tmp_path):
    """
    Verify that the CLI sanitizes the input filename stem before using it in output paths.
    """
    # Create a dummy input path that would cause traversal if not sanitized
    # Path("..").stem is ".."
    input_path = "path/to/.."
    output_dir = tmp_path / "output"
    output_dir.mkdir()

    args = MagicMock()
    args.input = input_path
    args.output_dir = str(output_dir)
    args.model = None
    args.api_key = None
    args.base_url = None
    args.timeout = 600.0
    args.rpm = None
    args.no_vision = True
    args.rag = True
    args.system_prompt = None
    args.dpi = None
    args.include_logic = False
    args.visual_summaries = False
    args.verbose = False

    # Mock extractor to avoid actual processing
    with patch("kami_excel_extractor.cli.KamiExcelExtractor") as mock_extractor_cls:
        mock_extractor = mock_extractor_cls.return_value
        # Mock aextract_rag_chunks as it's called when args.rag is True
        mock_extractor.aextract_rag_chunks = AsyncMock(return_value=({}, {}))

        # We also need to mock asyncio.to_thread because it's used for saving files
        with patch("kami_excel_extractor.cli.asyncio.to_thread", new_callable=AsyncMock) as mock_to_thread:
            await run_async(args)

            # Check the paths passed to asyncio.to_thread (which calls _save_json)
            # The first call should be for the result JSON
            # The second call should be for the RAG JSON

            called_paths = [call.args[1] for call in mock_to_thread.call_args_list]

            for path in called_paths:
                # Path should be in the output directory
                assert Path(path).parent == output_dir
                # Filename should not contain ".."
                assert ".." not in Path(path).name
                # Filename should be sanitized (e.g., "unnamed_result.json")
                assert "unnamed" in Path(path).name

def test_secure_filename_behavior():
    from kami_excel_extractor.utils import secure_filename
    assert secure_filename("..") == "unnamed"
    assert secure_filename(".") == "unnamed"
    assert secure_filename("/") == "unnamed"
    assert secure_filename("../etc/passwd") == "etc_passwd"
