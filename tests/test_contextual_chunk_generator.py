from datetime import datetime

import pytest

from kami_excel_extractor.rag_converter import ContextualChunkGenerator
from kami_excel_extractor.schema import RagOptions


def test_chunk_text_by_chars():
    options = RagOptions(max_chunk_chars=30, chunk_overlap_lines=1)
    generator = ContextualChunkGenerator(options=options)

    text = "line1: abcde\nline2: fghij\nline3: klmno\nline4: pqrst"
    chunks = generator._chunk_text_by_chars(text, max_chars=30, overlap_lines=1)

    assert len(chunks) >= 2
    assert "line1: abcde" in chunks[0]
    assert "line2: fghij" in chunks[1]


def test_find_coordinates_and_logic():
    options = RagOptions(include_logic=True, include_logic_annotations=True)
    generator = ContextualChunkGenerator(options=options)

    cells = [
        {"coord": "B2", "value": "売上合計", "formula": "=SUM(B3:B10)", "unit": "JPY"},
        {"coord": "B3", "value": 1500, "formula": None, "unit": None},
    ]

    section_text = "# 財務実績\n売上合計は好調です。"
    coord_range, has_formulas, annotations, formula_cells = generator._find_coordinates_and_logic(
        section_text, cells, include_logic_annotations=True
    )

    assert coord_range == "B2"
    assert has_formulas is True
    assert len(annotations) == 1
    assert "SUM(B3:B10)" in annotations[0]
    assert "JPY" in annotations[0]


def test_generate_chunks_yaml_frontmatter():
    options = RagOptions(output_format="yaml_frontmatter", max_chunk_chars=1000)
    generator = ContextualChunkGenerator(options=options)

    structured_content = {"sheets": {"Sheet1": {"Project": [{"ID": 1, "Name": "Alice"}]}}}

    raw_sheet_data = {
        "cells": [
            {"coord": "A1", "value": "Project", "formula": None},
            {"coord": "B1", "value": "Alice", "formula": None},
        ],
        "media": [],
    }

    chunks = generator.generate_chunks(
        sheet_name="Sheet1",
        structured_content=structured_content,
        raw_sheet_data=raw_sheet_data,
        source_file="test.xlsx",
    )

    assert len(chunks) == 1
    chunk = chunks[0]
    assert "content" in chunk
    assert "metadata" in chunk

    content = chunk["content"]
    assert "---" in content
    assert "source_file: test.xlsx" in content
    assert "sheet_name: Sheet1" in content
    assert "chunk_index: 1" in content
    assert "total_chunks: 1" in content
