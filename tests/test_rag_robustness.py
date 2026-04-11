import pytest
from kami_excel_extractor.rag_converter import RagChunker

def test_rag_chunker_complex_nesting():
    """深い見出し構造を持つ Markdown のチャンク化を検証"""
    md = """
# Section 1
Content 1
## Subsection 1.1
Content 1.1
# Section 2
Content 2
    """
    chunker = RagChunker(metadata={"file": "test.xlsx"})
    chunks = chunker.chunk(md, source_id="source_1")
    
    assert len(chunks) == 2
    assert chunks[0]["metadata"]["section"] == "Section 1"
    assert "Subsection 1.1" in chunks[0]["content"]
    assert chunks[1]["metadata"]["section"] == "Section 2"

def test_rag_chunker_empty_input():
    """空の入力に対する挙動"""
    chunker = RagChunker()
    assert chunker.chunk("") == []
    assert chunker.chunk("\n\n  \n") == []

def test_rag_chunker_no_headers():
    """見出しが全くない場合"""
    md = "Just some text without any headers."
    chunker = RagChunker()
    chunks = chunker.chunk(md)
    assert len(chunks) == 1
    assert chunks[0]["metadata"]["section"] == ""
    assert chunks[0]["content"] == md

def test_rag_chunker_with_formatting():
    """太字やリストが含まれる場合のチャンク化"""
    md = "# List Section\n- Item 1\n- **Bold Item 2**"
    chunker = RagChunker()
    chunks = chunker.chunk(md)
    assert len(chunks) == 1
    assert "- Item 1" in chunks[0]["content"]
    assert "**Bold Item 2**" in chunks[0]["content"]
