import asyncio
import pytest
from unittest.mock import patch, AsyncMock, MagicMock
from pathlib import Path
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.mark.asyncio
async def test_semaphore_limit_concurrency(monkeypatch):
    """Semaphore が並列実行数を正しく制限しているか検証"""
    # RPM 制限を 2 に設定
    monkeypatch.setenv("LLM_RPM_LIMIT", "2")
    extractor = KamiExcelExtractor(api_key="fake")
    
    current_concurrency = 0
    max_observed_concurrency = 0
    lock = asyncio.Lock()

    # モック関数内で semaphore を適切に利用する
    async def mock_aextract_single_sheet(name, content, model, system_prompt, image_url, semaphore, **kwargs):
        nonlocal current_concurrency, max_observed_concurrency
        
        # 呼び出し時に渡された semaphore を使ってロックをシミュレート
        async with semaphore:
            async with lock:
                current_concurrency += 1
                max_observed_concurrency = max(max_observed_concurrency, current_concurrency)
            
            # 処理時間をシミュレート
            await asyncio.sleep(0.1)
            
            async with lock:
                current_concurrency -= 1
        
        return name, {"data": "ok"}

    extractor._aextract_single_sheet = mock_aextract_single_sheet
    
    mock_raw_data = {
        "sheets": {
            "S1": "c", "S2": "c", "S3": "c", "S4": "c", "S5": "c"
        }
    }
    
    with patch("kami_excel_extractor.core.asyncio.to_thread") as mock_to_thread:
        mock_to_thread.side_effect = [mock_raw_data, None]
        await extractor.aextract_structured_data("dummy.xlsx")

    # 最大並列数が RPM 制限の 2 以下であることを確認
    assert max_observed_concurrency <= 2
    assert max_observed_concurrency > 0

@pytest.mark.asyncio
async def test_semaphore_default_limit(monkeypatch):
    """環境変数がない場合のデフォルトの Semaphore 動作を確認"""
    monkeypatch.delenv("LLM_RPM_LIMIT", raising=False)
    extractor = KamiExcelExtractor(api_key="fake")
    
    sem = extractor._get_semaphore()
    assert sem._value == 15
