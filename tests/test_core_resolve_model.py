import pytest
from kami_excel_extractor.core import KamiExcelExtractor

def test_resolve_model_default(monkeypatch):
    """Test that the default model is 'gemini-1.5-flash' when no env vars are set."""
    monkeypatch.delenv("LLM_MODEL", raising=False)
    monkeypatch.delenv("GEMINI_MODEL", raising=False)
    extractor = KamiExcelExtractor(api_key="fake")
    assert extractor._resolve_model() == "gemini-1.5-flash"

def test_resolve_model_llm_model(monkeypatch):
    """Test that LLM_MODEL environment variable is used."""
    monkeypatch.setenv("LLM_MODEL", "gpt-4o")
    monkeypatch.delenv("GEMINI_MODEL", raising=False)
    extractor = KamiExcelExtractor(api_key="fake")
    assert extractor._resolve_model() == "gpt-4o"

def test_resolve_model_gemini_model(monkeypatch):
    """Test that GEMINI_MODEL environment variable is used as fallback."""
    monkeypatch.delenv("LLM_MODEL", raising=False)
    monkeypatch.setenv("GEMINI_MODEL", "gemini-2.0-flash")
    extractor = KamiExcelExtractor(api_key="fake")
    assert extractor._resolve_model() == "gemini-2.0-flash"

def test_resolve_model_precedence(monkeypatch):
    """Test that LLM_MODEL takes precedence over GEMINI_MODEL."""
    monkeypatch.setenv("LLM_MODEL", "gpt-4o")
    monkeypatch.setenv("GEMINI_MODEL", "gemini-2.0-flash")
    extractor = KamiExcelExtractor(api_key="fake")
    assert extractor._resolve_model() == "gpt-4o"

def test_resolve_model_override(monkeypatch):
    """Test that explicit model argument overrides the default."""
    monkeypatch.delenv("LLM_MODEL", raising=False)
    monkeypatch.delenv("GEMINI_MODEL", raising=False)
    extractor = KamiExcelExtractor(api_key="fake")
    assert extractor._resolve_model("claude-3-opus") == "claude-3-opus"

def test_resolve_model_fallback_none(monkeypatch):
    """Test that passing None to _resolve_model falls back to the default."""
    monkeypatch.setenv("LLM_MODEL", "gpt-4o")
    extractor = KamiExcelExtractor(api_key="fake")
    assert extractor._resolve_model(None) == "gpt-4o"
