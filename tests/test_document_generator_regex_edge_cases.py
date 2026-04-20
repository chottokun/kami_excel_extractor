
import pytest
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_image_regex_balanced_parentheses_levels(doc_gen):
    """括弧が2重以上にネストされた場合の挙動を検証 (現在の正規表現は1レベルまでを想定)"""
    # 1レベル（サポート対象）
    md = "![alt](image(1).png)"
    html = doc_gen._simple_md_to_html(md)
    assert 'src="image(1).png"' in html
    
    # 2レベル（新しい手動パースにより、任意のレベルのネストがサポートされる）
    md2 = "![alt](image((nested)).png)"
    html2 = doc_gen._simple_md_to_html(md2)
    assert 'src="image((nested)).png"' in html2

def test_image_regex_unbalanced_parentheses(doc_gen):
    """括弧が閉じられていない、または余計な括弧がある場合の挙動"""
    # 閉じが足りない -> マッチしないはず
    md = "![alt](image(1.png)"
    html = doc_gen._simple_md_to_html(md)
    assert '<img' not in html
    assert '![alt](image(1.png)' in html # そのまま出力される

    # 開始が足りない -> 最初の ) でマッチが終了する
    md2 = "![alt](image)1.png)"
    html2 = doc_gen._simple_md_to_html(md2)
    # 現在の正規表現では `(image)` までがパスとして認識される
    assert 'src="image"' in html2

def test_image_regex_special_characters_in_path(doc_gen):
    """パスに特殊文字や日本語、スペースが含まれる場合"""
    md = "![alt](画像 (2024).png?query=1&param=2)"
    html = doc_gen._simple_md_to_html(md)
    # HTMLエスケープにより & -> &amp; になる
    assert 'src="画像 (2024).png?query=1&amp;param=2"' in html

def test_multiple_images_on_same_line_limitation(doc_gen):
    """1行に複数の画像がある場合、現在の実装では最初の1つしか処理されないことを確認(ドキュメント化された制限)"""
    md = "![img1](path1.png) ![img2](path2.png)"
    html = doc_gen._simple_md_to_html(md)
    assert 'src="path1.png"' in html
    assert 'src="path2.png"' not in html
    # 1つの div の中に1つの img だけが入る
    assert html.count('<img') == 1

def test_image_at_end_of_line(doc_gen):
    """行の途中に画像がある場合（!で始まらない場合）の挙動"""
    md = "ここに画像があります: ![img](path.png)"
    html = doc_gen._simple_md_to_html(md)
    # _simple_md_to_html の実装を見ると `if stripped.startswith('!['):` となっているため
    # 行頭にない画像は現在はパースされないはず
    assert '<img' not in html
