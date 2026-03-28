import pytest
import os
from kami_excel_extractor.core import KamiExcelExtractor

def test_resolve_model_universal():
    # LLM_MODEL が最優先されることを確認
    os.environ["LLM_MODEL"] = "anthropic/claude-3"
    os.environ["GEMINI_MODEL"] = "gemini-1.5-flash"
    extractor = KamiExcelExtractor()
    
    # 引数指定なしの場合、LLM_MODEL が使われる (プレフィックスはそのまま)
    assert extractor._resolve_model(None) == "anthropic/claude-3"
    
    # 引数で直接指定した場合は、それが最優先 (プレフィックス制限なし)
    assert extractor._resolve_model("openai/gpt-4o") == "openai/gpt-4o"
    assert extractor._resolve_model("ollama/qwen3.5:4b") == "ollama/qwen3.5:4b"

def test_resolve_model_fallback():
    # LLM_MODEL がなく、GEMINI_MODEL がある場合
    if "LLM_MODEL" in os.environ:
        del os.environ["LLM_MODEL"]
    os.environ["GEMINI_MODEL"] = "gemini-1.5-pro"
    extractor = KamiExcelExtractor()
    
    # 既存の fallback 動作の確認
    assert extractor._resolve_model(None) == "gemini-1.5-pro"

def test_llm_config_passing():
    # ベースURLとタイムアウトの引き継ぎ確認
    extractor = KamiExcelExtractor(
        base_url="http://my-ollama:11434",
        timeout=120.0
    )
    assert extractor.base_url == "http://my-ollama:11434"
    assert extractor.timeout == 120.0

def test_rpm_limit_from_env():
    # RPM制限が環境変数から読み取られるか
    os.environ["LLM_RPM_LIMIT"] = "50"
    extractor = KamiExcelExtractor()
    sem = extractor._get_semaphore()
    # セマフォの内部値を確認 (litellm_rpm_limit に基づく)
    assert extractor.litellm_rpm_limit == 50
