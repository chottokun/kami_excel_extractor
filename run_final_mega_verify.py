
import asyncio
import json
from pathlib import Path
from unittest.mock import patch, AsyncMock
from kami_excel_extractor.core import KamiExcelExtractor, ExtractionOptions

async def test_mega_extraction():
    extractor = KamiExcelExtractor(output_dir="mega_test_out")
    excel_path = Path("tests/assets/mega_mixed_report.xlsx")

    print("\n--- [Final Verification] Mega Sample Extraction ---")
    with patch("litellm.acompletion") as mock_completion:
        mock_completion.return_value.choices = [
            AsyncMock(message=AsyncMock(content='```json\n{"data": [], "summary": "Success"}\n```'))
        ]
        
        # 1. キャッシュありでの実行
        options = ExtractionOptions(include_visual_summaries=True, include_logic=True, use_cache=True)
        print("Run 1 (Caching enabled)...")
        await extractor.aextract_structured_data(excel_path, options=options)
        print(f"API Calls in Run 1: {mock_completion.call_count}")
        
        # 2. キャッシュありでの再実行 (APIが呼ばれないことを期待)
        mock_completion.reset_mock()
        print("Run 2 (Caching enabled, should hit cache)...")
        await extractor.aextract_structured_data(excel_path, options=options)
        print(f"API Calls in Run 2: {mock_completion.call_count}")
        
        # 3. キャッシュ無効での再実行 (APIが呼ばれることを期待)
        mock_completion.reset_mock()
        print("Run 3 (Cache disabled, should bypass cache)...")
        options.use_cache = False
        await extractor.aextract_structured_data(excel_path, options=options)
        print(f"API Calls in Run 3: {mock_completion.call_count}")

    print("\n✅ Verification complete. Check mega_test_out for artifacts.")

if __name__ == "__main__":
    asyncio.run(test_mega_extraction())
