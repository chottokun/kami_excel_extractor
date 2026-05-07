
import pytest
from pathlib import Path
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

@pytest.mark.parametrize("line,expected", [
    ("![alt](path.png)", ("alt", "path.png")),
    ("![alt](image(1).png)", ("alt", "image(1).png")),
    ("![alt](image((nested)).png)", ("alt", "image((nested)).png")),
    ("![](path.png)", ("", "path.png")),
    ("![my image](path/to my/image (1).png)", ("my image", "path/to my/image (1).png")),
    ("![alt](path-with-chars-@#$.png)", ("alt", "path-with-chars-@#$.png")),
    ("  ![alt](path.png)  ", ("alt", "path.png")),
    ("![alt](path)trailing", ("alt", "path")),
    ("![alt](path(nested)with)multiple(parentheses)", ("alt", "path(nested)with")),
])
def test_parse_balanced_image_happy_paths(doc_gen, line, expected):
    assert doc_gen._parse_balanced_image(line) == expected

@pytest.mark.parametrize("line", [
    "[alt](path)",             # Missing !
    "![alt]path",              # Missing ()
    "![alt](path",             # Unbalanced (missing closing)
    "![alt] (path)",           # Space between ] and (
    "random text",             # Just text
    "![alt]",                  # Missing path part entirely
    "![alt](",                 # Missing path content and closing
    "! [alt](path)",           # Space between ! and [
    "![alt](path(unbalanced)", # Unbalanced nested
])
def test_parse_balanced_image_invalid_formats(doc_gen, line):
    assert doc_gen._parse_balanced_image(line) is None

def test_parse_balanced_image_complex_nesting(doc_gen):
    line = "![alt](path(a(b)c)d)"
    assert doc_gen._parse_balanced_image(line) == ("alt", "path(a(b)c)d")
