import logging
import os
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock, AsyncMock, mock_open
import pytest
from kami_excel_extractor.cli import main
from kami_excel_extractor.schema import ExtractionOptions, RagOptions

@pytest.fixture
def mock_extractor():
    with patch("kami_excel_extractor.cli.KamiExcelExtractor") as mock:
        yield mock

def test_cli_new_flags_passing(mock_extractor):
    """新しく追加されたフラグ (--include-logic, --visual-summaries) が正しく渡されることを確認"""
    mock_inst = mock_extractor.return_value
    mock_inst.aextract_structured_data = AsyncMock()
    mock_inst.aextract_structured_data.return_value = {}

    test_args = [
        "kami-excel", "test.xlsx",
        "--include-logic",
        "--visual-summaries",
        "--dpi", "300"
    ]

    m_open = mock_open()

    with patch("sys.argv", test_args), \
         patch("pathlib.Path.exists", return_value=True), \
         patch("builtins.open", m_open), \
         patch("json.dump"):

        main()

        # aextract_structured_data の引数に新フラグが反映されているか確認
        mock_inst.aextract_structured_data.assert_called_once_with(
            "test.xlsx",
            options=ExtractionOptions(
                model=None,
                system_prompt=None,
                include_visual_summaries=True,
                include_logic=True,
                use_visual_context=True,
                dpi=300
            )
        )

def test_cli_rag_new_flags_passing(mock_extractor):
    """RAGモードでも新フラグが正しく渡されることを確認"""
    mock_inst = mock_extractor.return_value
    mock_inst.aextract_rag_chunks = AsyncMock()
    mock_inst.aextract_rag_chunks.return_value = ({}, {})

    test_args = [
        "kami-excel", "test.xlsx",
        "--rag",
        "--include-logic",
        "--visual-summaries"
    ]

    m_open = mock_open()

    with patch("sys.argv", test_args), \
         patch("pathlib.Path.exists", return_value=True), \
         patch("builtins.open", m_open), \
         patch("json.dump"):

        main()

        # aextract_rag_chunks の引数を確認
        mock_inst.aextract_rag_chunks.assert_called_once_with(
            "test.xlsx",
            options=RagOptions(
                model=None,
                include_logic=True,
                include_visual_summaries=True,
                use_visual_context=True,
                dpi=150
            )
        )

def test_cli_visual_summaries_explicit_false(mock_extractor):
    """--no-vision 時には visual-summaries がデフォルトで False になることを確認"""
    mock_inst = mock_extractor.return_value
    mock_inst.aextract_structured_data = AsyncMock()
    mock_inst.aextract_structured_data.return_value = {}

    test_args = [
        "kami-excel", "test.xlsx",
        "--no-vision"
    ]

    m_open = mock_open()

    with patch("sys.argv", test_args), \
         patch("pathlib.Path.exists", return_value=True), \
         patch("builtins.open", m_open), \
         patch("json.dump"):

        main()

        # include_visual_summaries が False であることを確認
        mock_inst.aextract_structured_data.assert_called_once_with(
            "test.xlsx",
            options=ExtractionOptions(
                model=None,
                system_prompt=None,
                include_visual_summaries=False,
                include_logic=False,
                use_visual_context=False,
                dpi=150
            )
        )
