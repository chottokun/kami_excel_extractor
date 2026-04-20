
import pytest
import asyncio
from unittest.mock import patch, AsyncMock, MagicMock
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(tmp_path):
    return KamiExcelExtractor(output_dir=tmp_path)

@pytest.mark.asyncio
async def test_awith_retry_success_first_time(extractor):
    """1回目で成功する場合、リトライされないことを検証"""
    mock_func = AsyncMock(return_value="success")
    
    result = await extractor._awith_retry(mock_func)
    
    assert result == "success"
    assert mock_func.call_count == 1

@pytest.mark.asyncio
async def test_awith_retry_exponential_backoff(extractor):
    """レートリミット(429)時にリトライされ、最終的に成功することを検証"""
    mock_func = AsyncMock()
    
    # 2回失敗（429）、3回目で成功
    error = Exception("Rate limit")
    error.status_code = 429
    mock_func.side_effect = [error, error, "success"]
    
    # テスト時間を短縮するために initial_delay を小さく設定
    with patch("asyncio.sleep", AsyncMock()) as mock_sleep:
        result = await extractor._awith_retry(mock_func, initial_delay=0.01)
        
        assert result == "success"
        assert mock_func.call_count == 3
        # sleepが2回呼ばれていることを確認 (1回目: 0.01s, 2回目: 0.02s)
        assert mock_sleep.call_count == 2

@pytest.mark.asyncio
async def test_awith_retry_max_retries_exceeded(extractor):
    """最大リトライ回数を超えても失敗し続ける場合、最後の例外を投げることを検証"""
    mock_func = AsyncMock()
    error = Exception("Rate limit")
    error.status_code = 429
    mock_func.side_effect = error
    
    with patch("asyncio.sleep", AsyncMock()):
        with pytest.raises(Exception) as excinfo:
            await extractor._awith_retry(mock_func, max_retries=2, initial_delay=0.01)
        
        assert str(excinfo.value) == "Rate limit"
        assert mock_func.call_count == 3 # 初回 + リトライ2回

@pytest.mark.asyncio
async def test_awith_retry_non_retryable_error(extractor):
    """リトライ対象外のエラー（例：401 Unauthorized）では即座に失敗することを検証"""
    mock_func = AsyncMock()
    error = Exception("Unauthorized")
    error.status_code = 401
    mock_func.side_effect = error
    
    with patch("asyncio.sleep", AsyncMock()) as mock_sleep:
        with pytest.raises(Exception) as excinfo:
            await extractor._awith_retry(mock_func)
            
        assert str(excinfo.value) == "Unauthorized"
        assert mock_func.call_count == 1
        assert mock_sleep.call_count == 0
