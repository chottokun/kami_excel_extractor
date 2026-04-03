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
    result = converter.convert(data)
    assert "## Project" in result
    assert "- **ID**: 123" in result
    assert "- A" in result
    assert "- B" in result

def test_json_to_markdown_table_success():
    converter = JsonToMarkdownConverter()
    data = [
        {"ID": 1, "Name": "Alice"},
        {"ID": 2, "Name": "Bob"}
    ]
    result = converter.convert(data)
    assert "| ID | Name |" in result
    assert "| 1 | Alice |" in result
    assert "| 2 | Bob |" in result

def test_json_to_markdown_kv_format():
    """list_format='kv' の場合にリストが KV 形式で出力されることを確認"""
    converter = JsonToMarkdownConverter(list_format="kv")
    data = [
        {"ID": 1, "Name": "Alice"},
        {"ID": 2, "Name": "Bob"}
    ]
    result = converter.convert(data)
    assert "- ID: 1, Name: Alice" in result
    assert "- ID: 2, Name: Bob" in result
    assert "|" not in result

def test_json_to_markdown_table_inconsistent_keys():
    """キーが一致しないリストはテーブル化されず、箇条書きになることを確認"""
    converter = JsonToMarkdownConverter()
    data = [
        {"ID": 1, "Name": "Alice"},
        {"ID": 2, "Age": 30}
    ]
    result = converter.convert(data)
    assert "|" not in result
    assert "- **ID**: 1" in result
    assert "- **Age**: 30" in result

def test_json_to_markdown_empty_data():
    converter = JsonToMarkdownConverter()
    assert converter.convert({}) == ""
    assert converter.convert([]) == ""
    assert converter.convert(None) == "None"

def test_json_to_markdown_media():
    converter = JsonToMarkdownConverter()
    data = {
        "media": [
            {"filename": "img1.png", "visual_summary": "[画像概要] グラフの要約"},
            {"filename": "img2.png", "visual_summary": "概要prefixなし"}
        ]
    }
    result = converter.convert(data)
    assert "## 関連メディア" in result
    assert "![画像](media/img1.png)" in result
    assert "**[画像概要]**: グラフの要約" in result
    assert "**[画像概要]**: 概要prefixなし" in result

def test_rag_chunker_basic():
    chunker = RagChunker(metadata={"file": "test.xlsx"})
    markdown = "# Section 1\nContent A\n# Section 2\nContent B"
    chunks = chunker.chunk(markdown, "test.xlsx")
    
    assert len(chunks) == 2
    assert chunks[0]["metadata"]["section"] == "Section 1"
    assert "Content A" in chunks[0]["content"]
    assert chunks[1]["metadata"]["section"] == "Section 2"
    assert "Content B" in chunks[1]["content"]
    assert chunks[0]["metadata"]["file"] == "test.xlsx"

def test_rag_chunker_no_headers():
    chunker = RagChunker()
    markdown = "Just plain text\nNo headers here"
    chunks = chunker.chunk(markdown, "test.xlsx")
    assert len(chunks) == 1
    assert chunks[0]["metadata"]["section"] == ""
    assert "Just plain text" in chunks[0]["content"]

def test_rag_chunker_complex_headers():
    chunker = RagChunker()
    markdown = """# Title
Intro
## Subtitle
Details
# Another Section
End"""
    chunks = chunker.chunk(markdown, "test.xlsx")
    # 現状の実装は "# " (トップレベル) のみで分割
    assert len(chunks) == 2
    assert chunks[0]["metadata"]["section"] == "Title"
    assert "## Subtitle" in chunks[0]["content"]
    assert chunks[1]["metadata"]["section"] == "Another Section"

def test_json_to_markdown_list_strings():
    """Verify that a list of strings is converted to a markdown list"""
    converter = JsonToMarkdownConverter()
    data = ["Apple", "Banana", "Cherry"]
    result = converter.convert(data)
    expected = "- Apple\n- Banana\n- Cherry"
    assert result == expected

def test_json_to_markdown_list_integers():
    """Verify that a list of integers is converted to a markdown list"""
    converter = JsonToMarkdownConverter()
    data = [1, 2, 3]
    result = converter.convert(data)
    expected = "- 1\n- 2\n- 3"
    assert result == expected

def test_json_to_markdown_list_mixed():
    """Verify that a list of mixed types is correctly processed"""
    converter = JsonToMarkdownConverter()
    data = ["Text", 123, None]
    result = converter.convert(data)
    expected = "- Text\n- 123\n- None"
    assert result == expected

def test_json_to_markdown_list_nested():
    """Verify the behavior for nested lists.

    Note: The current implementation produces suboptimal markdown for nested lists
    (sub-items are not properly indented under the parent item), but these tests
    ensure that the current behavior is preserved and documented.
    """
    converter = JsonToMarkdownConverter()
    data = ["Outer", ["Inner1", "Inner2"]]
    result = converter.convert(data)
    # The current implementation (level management) results in:
    # - Outer
    # - - Inner1
    # - Inner2
    expected = "- Outer\n- - Inner1\n- Inner2"
    assert result == expected
