import pytest
import asyncio
from unittest.mock import patch, MagicMock
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.mark.asyncio
async def test_semaphore_limit_concurrency(monkeypatch):
    """Semaphore が並列実行数を正しく制限しているか検証"""
    # 明示的に RPM 制限を 2 に設定
    extractor = KamiExcelExtractor(api_key="fake", litellm_rpm_limit=2)

    current_concurrency = 0
    max_observed_concurrency = 0
    lock = asyncio.Lock()

    mock_raw_data = {
        "sheets": {
            "S1": {"is_simple": False}, "S2": {"is_simple": False}, 
            "S3": {"is_simple": False}, "S4": {"is_simple": False}, "S5": {"is_simple": False}
        }
    }

    # クラスレベルで確実にパッチ
    with patch.object(KamiExcelExtractor, "_aextract_single_sheet", autospec=True) as mock_meth, \
         patch("kami_excel_extractor.core.asyncio.to_thread", return_value=mock_raw_data), \
         patch("kami_excel_extractor.core.ExcelConverter.convert", return_value="dummy.png"):
        
        # モックの挙動を定義
        async def side_effect(self, name, content, model, system_prompt, image_url, semaphore, **kwargs):
            nonlocal current_concurrency, max_observed_concurrency
            async with semaphore:
                async with lock:
                    current_concurrency += 1
                    max_observed_concurrency = max(max_observed_concurrency, current_concurrency)
                await asyncio.sleep(0.1)
                async with lock:
                    current_concurrency -= 1
            return name, {"data": "ok"}
        
        mock_meth.side_effect = side_effect
        
        await extractor.aextract_structured_data("dummy.xlsx")

    # 最大並列数が RPM 制限の 2 以下であることを確認
    assert max_observed_concurrency <= 2
    assert max_observed_concurrency > 0

@pytest.mark.asyncio
async def test_semaphore_default_limit(monkeypatch):
    """環境変数がない場合のデフォルトの Semaphore 動作を確認"""
    monkeypatch.delenv("LLM_RPM_LIMIT", raising=False)
    # 明示的に RPM 制限を指定しないインスタンスを作成
    extractor = KamiExcelExtractor(api_key="fake", litellm_rpm_limit=0)
    
    sem = extractor._get_semaphore()
    # 新しいデフォルト値は 1000
    assert sem._value == 1000
