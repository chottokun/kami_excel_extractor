import re
import unicodedata

# Compiled regex patterns for performance
_FILENAME_SANITIZE_RE = re.compile(r'[^\w\.\-]')
_FILENAME_MULTIPLE_UNDERSCORES_RE = re.compile(r'__+')
_FILENAME_MULTIPLE_DOTS_RE = re.compile(r'\.\.+')

def clean_kami_text(text: any) -> any:
    """
    Excel内のテキストをクリーニングする。
    - 全角スペースを半角スペースに正規化
    - 漢字・ひらがな・カタカナの間に挟まった不自然な空白(1-3個)を削除 (例: "氏  名" -> "氏名")
    """
    if not isinstance(text, str):
        return text

    # 全角スペースを半角に
    text = text.replace('\u3000', ' ')
    
    # 漢字・ひらがな・カタカナの間に挟まった1〜3つの空白を削除
    # 正規表現の解説:
    # ([\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff]) : 前方の文字 (和字)
    # \s{1,3} : 1〜3個の空白
    # (?=[\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff]) : 後方の文字 (和字) を先読み
    res = re.sub(r'([\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])\s{1,3}(?=[\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])', r'\1', text)
    
    return res.strip()

def secure_filename(filename: str) -> str:
    """
    Sanitize a string to be used as a filename.
    Keeps only alphanumeric characters, underscores, dashes, and dots.
    Converts spaces to underscores and removes path separators.
    """
    if not filename:
        return "unnamed"

    # Normalize unicode characters to NFKD (separate base characters from marks)
    filename = unicodedata.normalize('NFKD', filename)

    # Replace spaces with underscores
    filename = filename.replace(" ", "_")

    # Remove anything that isn't alphanumeric, underscore, dash, or dot
    # We allow Japanese/other characters if they are alphanumeric according to unicode
    # But for maximum security, we might want to restrict to ASCII if we don't care about Japanese names.
    # However, this tool works with Japanese Excel files, so we should support Japanese.

    # Keep alphanumeric (including Japanese), underscores, dashes, and dots.
    # [^\w\.\-] where \w includes Unicode word characters.
    filename = _FILENAME_SANITIZE_RE.sub('_', filename)

    # Strip leading/trailing dots and underscores
    filename = filename.strip("._")

    # Remove multiple consecutive underscores/dots
    filename = _FILENAME_MULTIPLE_UNDERSCORES_RE.sub('_', filename)
    filename = _FILENAME_MULTIPLE_DOTS_RE.sub('.', filename)

    # If the filename becomes empty, use a default
    if not filename or filename in (".", ".."):
        return "unnamed"

    return filename
