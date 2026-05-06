import re
import unicodedata
import sqlite3
import hashlib
from pathlib import Path
from typing import Optional, Any

# Compiled regex patterns for performance
_FILENAME_SANITIZE_RE = re.compile(r'[^\w\.\-]')
_FILENAME_MULTIPLE_UNDERSCORES_RE = re.compile(r'__+')
_FILENAME_MULTIPLE_DOTS_RE = re.compile(r'\.\.+')

def clean_kami_text(text: Any) -> Any:
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

    filename = unicodedata.normalize('NFKD', filename)
    filename = filename.replace(" ", "_")
    filename = _FILENAME_SANITIZE_RE.sub('_', filename)
    filename = filename.strip("._")
    filename = _FILENAME_MULTIPLE_UNDERSCORES_RE.sub('_', filename)
    filename = _FILENAME_MULTIPLE_DOTS_RE.sub('.', filename)

    if not filename or filename in (".", ".."):
        return "unnamed"

    return filename

class CacheManager:
    """SQLiteを使用したキャッシュ永続化マネージャー"""
    
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        """データベースとテーブルの初期化"""
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS vlm_cache (
                    key TEXT PRIMARY KEY,
                    content TEXT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS image_cache (
                    hash TEXT PRIMARY KEY,
                    data_url TEXT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS llm_cache (
                    key TEXT PRIMARY KEY,
                    content TEXT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)

    def get_file_hash(self, file_path: Path) -> str:
        """ファイルの内容からSHA-256ハッシュを生成する"""
        sha256 = hashlib.sha256()
        with open(file_path, "rb") as f:
            while chunk := f.read(8192):
                sha256.update(chunk)
        return sha256.hexdigest()

    def get_vlm_result(self, model: str, prompt: str, image_hash: str) -> Optional[str]:
        """VLMの解析結果をキャッシュから取得"""
        key = f"{model}:{prompt}:{image_hash}"
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute("SELECT content FROM vlm_cache WHERE key = ?", (key,))
            row = cur.fetchone()
            return row[0] if row else None

    def set_vlm_result(self, model: str, prompt: str, image_hash: str, content: str):
        """VLMの解析結果をキャッシュに保存"""
        key = f"{model}:{prompt}:{image_hash}"
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR REPLACE INTO vlm_cache (key, content) VALUES (?, ?)", (key, content))

    def get_image_data_url(self, image_hash: str) -> Optional[str]:
        """Base64データURLをキャッシュから取得"""
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute("SELECT data_url FROM image_cache WHERE hash = ?", (image_hash,))
            row = cur.fetchone()
            return row[0] if row else None

    def set_image_data_url(self, image_hash: str, data_url: str):
        """Base64データURLをキャッシュに保存"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR REPLACE INTO image_cache (hash, data_url) VALUES (?, ?)", (image_hash, data_url))

    def get_llm_result(self, model: str, prompt: str, input_text: str) -> Optional[str]:
        """LLMの解析結果をキャッシュから取得"""
        input_hash = hashlib.sha256(input_text.encode("utf-8")).hexdigest()
        key = f"{model}:{hashlib.sha256(prompt.encode('utf-8')).hexdigest()}:{input_hash}"
        with sqlite3.connect(self.db_path) as conn:
            cur = conn.execute("SELECT content FROM llm_cache WHERE key = ?", (key,))
            row = cur.fetchone()
            return row[0] if row else None

    def set_llm_result(self, model: str, prompt: str, input_text: str, content: str):
        """LLMの解析結果をキャッシュに保存"""
        input_hash = hashlib.sha256(input_text.encode("utf-8")).hexdigest()
        key = f"{model}:{hashlib.sha256(prompt.encode('utf-8')).hexdigest()}:{input_hash}"
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("INSERT OR REPLACE INTO llm_cache (key, content) VALUES (?, ?)", (key, content))

    def clear(self):
        """すべてのキャッシュを削除する"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM vlm_cache")
            conn.execute("DELETE FROM image_cache")
            conn.execute("DELETE FROM llm_cache")
            conn.commit()
