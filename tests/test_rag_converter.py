import pytest
from kami_excel_extractor.rag_converter import JsonToMarkdownConverter, RagChunker

def test_json_to_markdown_simple():
    converter = JsonToMarkdownConverter()
    data = {"name": "Test Service", "status": "active"}
    expected = "- **name**: Test Service\n- **status**: active"
    assert converter.convert(data) == expected

def test_json_to_markdown_nested():
    converter = JsonToMarkdownConverter()
    data = {
        "Project": {
            "ID": 123,
            "Details": ["A", "B"]
        }
    }
    # level 1 dict -> # Project
    # ID/Details -> level 2
    result = converter.convert(data)
    assert "## Project" in result
    assert "- **ID**: 123" in result
    assert "- A" in result
    assert "- B" in result

def test_json_to_markdown_table():
    converter = JsonToMarkdownConverter()
    data = [
        {"ID": 1, "Name": "Alice"},
        {"ID": 2, "Name": "Bob"}
    ]
    result = converter.convert(data)
    assert "| ID | Name |" in result
    assert "| 1 | Alice |" in result
    assert "| 2 | Bob |" in result

def test_rag_chunker():
    chunker = RagChunker(metadata={"file": "test.xlsx"})
    markdown = "# Section 1\nContent A\n# Section 2\nContent B"
    chunks = chunker.chunk(markdown, "test.xlsx")
    
    assert len(chunks) == 2
    assert chunks[0]["metadata"]["section"] == "Section 1"
    assert "Content A" in chunks[0]["content"]
    assert chunks[1]["metadata"]["section"] == "Section 2"
    assert "Content B" in chunks[1]["content"]
    assert chunks[0]["metadata"]["file"] == "test.xlsx"
