import time
import os
import base64
from unittest.mock import MagicMock, patch
from pathlib import Path
from kami_excel_extractor.core import KamiExcelExtractor

def benchmark(latency=0.1):
    # Setup
    extractor = KamiExcelExtractor(api_key="fake")

    # Mock dependencies
    with patch("kami_excel_extractor.core.ExcelConverter.convert") as mock_convert, \
         patch("kami_excel_extractor.core.MetadataExtractor.extract") as mock_extract, \
         patch("litellm.completion") as mock_completion, \
         patch("builtins.open") as mock_open, \
         patch("pathlib.Path.exists") as mock_exists:

        mock_convert.return_value = Path("dummy.png")
        mock_exists.return_value = True # For media files

        # Mock file read for base64 encoding
        mock_file = MagicMock()
        mock_file.__enter__.return_value.read.return_value = b"fake binary"
        mock_open.return_value = mock_file

        # 5 sheets, each with 2 media items (total 15 LLM calls)
        num_sheets = 5
        num_media_per_sheet = 2

        sheets_data = {}
        for i in range(num_sheets):
            sheet_name = f"Sheet{i}"
            media = [{"filename": f"{sheet_name}_media_{j}.png"} for j in range(num_media_per_sheet)]
            sheets_data[sheet_name] = {"html": f"data{i}", "media": media}

        mock_extract.return_value = {"sheets": sheets_data}

        def slow_completion(*args, **kwargs):
            time.sleep(latency) # Simulate latency
            m = MagicMock()
            m.choices[0].message.content = "```yaml\nkey: value\n```"
            return m

        mock_completion.side_effect = slow_completion

        # Set RPM limit to 60 -> 1 second sleep between requests
        os.environ["GEMINI_RPM_LIMIT"] = "60"

        print(f"Starting benchmark with {num_sheets} sheets and {num_sheets * num_media_per_sheet} media items...")
        print(f"Latency: {latency}s, GEMINI_RPM_LIMIT=60 (1s sleep between requests)")

        start_time = time.time()
        extractor.extract_structured_data("dummy.xlsx", include_visual_summaries=True)
        end_time = time.time()

        duration = end_time - start_time
        print(f"Total time: {duration:.2f} seconds")

if __name__ == "__main__":
    print("--- Low Latency Scenario (0.1s) ---")
    benchmark(latency=0.1)
    print("\n--- High Latency Scenario (2.0s) ---")
    benchmark(latency=2.0)
