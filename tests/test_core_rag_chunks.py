import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch, AsyncMock
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions, RagOptions

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("kami_excel_extractor.core.KamiExcelExtractor.aextract_structured_data", new_callable=AsyncMock)
async def test_extract_rag_chunks_custom_format_and_default_model(mock_aextract_struct, mock_extract_raw, extractor, sample_excel_path):
    # Setup mocks
    mock_extract_raw.return_value = {"sheets": {"Sheet1": {"is_simple": False}}}

    # Simulate structured data return
    structured_data = {
        "sheets": {
            "Sheet1": {
                "Project": [
                    {"ID": 1, "Name": "Alice"},
                    {"ID": 2, "Name": "Bob"}
                ]
            }
        }
    }
    mock_aextract_struct.return_value = structured_data

    # Default format is 'table' in JsonToMarkdownConverter, but KamiExcelExtractor might change it.
    # In KamiExcelExtractor.__init__, self.rag_converter = JsonToMarkdownConverter() (default list_format='table')
    # Let's ensure we are testing the override logic.

    # Case 1: Override to 'kv' and use default model
    rag_options = RagOptions(model=None, list_format="kv")
    sheet_results, raw_data = await extractor.aextract_rag_chunks(sample_excel_path, options=rag_options)

    # Verification
    # 1. Default model used
    mock_aextract_struct.assert_called_with(
        Path(sample_excel_path),
        options=ExtractionOptions(
            model=None,
            system_prompt=None,
            include_visual_summaries=True,
            use_visual_context=True
        )
    )
    # The actual code in aextract_rag_chunks calls:
    # structured_data = await self.aextract_structured_data(excel_path, model=model, include_visual_summaries=True)
    # where model=None (passed from extract_rag_chunks)

    # 2. KV format used in markdown
    markdown = sheet_results["Sheet1"]["markdown"]
    assert "ID: 1, Name: Alice" in markdown
    assert "|" not in markdown

    # 3. List format restored (the code restores it in a finally block)
    assert extractor.rag_converter.list_format == "table"

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("kami_excel_extractor.core.KamiExcelExtractor.aextract_structured_data", new_callable=AsyncMock)
async def test_extract_rag_chunks_table_format(mock_aextract_struct, mock_extract_raw, extractor, sample_excel_path):
    mock_extract_raw.return_value = {"sheets": {"Sheet1": {"is_simple": False}}}
    structured_data = {
        "sheets": {
            "Sheet1": {
                "Project": [
                    {"ID": 1, "Name": "Alice"},
                    {"ID": 2, "Name": "Bob"}
                ]
            }
        }
    }
    mock_aextract_struct.return_value = structured_data

    # Explicitly set to table
    rag_options = RagOptions(model="openai/gpt-4o", list_format="table")
    sheet_results, _ = await extractor.aextract_rag_chunks(sample_excel_path, options=rag_options)

    markdown = sheet_results["Sheet1"]["markdown"]
    assert "| ID | Name |" in markdown
    assert "| 1 | Alice |" in markdown

    # Model passed correctly
    mock_aextract_struct.assert_called_with(
        Path(sample_excel_path),
        options=ExtractionOptions(
            model="openai/gpt-4o",
            system_prompt=None,
            include_visual_summaries=True,
            use_visual_context=True
        )
    )

def test_extract_rag_chunks_sync_wrapper(output_dir, sample_excel_path):
    # This tests the sync wrapper specifically to ensure it calls asyncio.run correctly
    with patch("kami_excel_extractor.core.KamiExcelExtractor.aextract_rag_chunks", new_callable=AsyncMock) as mock_aextract:
        mock_aextract.return_value = ({}, {})
        extractor = KamiExcelExtractor(api_key="fake", output_dir=str(output_dir))
        rag_options = RagOptions(model="test-model", list_format="kv")
        extractor.extract_rag_chunks(sample_excel_path, options=rag_options)
        mock_aextract.assert_called_once()
