import asyncio
import hashlib
import re
import sqlite3
import threading
import unicodedata
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Generator, Optional

# Compiled regex patterns for performance
_FILENAME_SANITIZE_RE = re.compile(r"[^\w\.\-]")
_FILENAME_MULTIPLE_UNDERSCORES_RE = re.compile(r"__+")
_FILENAME_MULTIPLE_DOTS_RE = re.compile(r"\.\.+")
_CLEAN_KAMI_TEXT_RE = re.compile(
    r"([\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])\s{1,3}(?=[\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])"
)
_JAPANESE_CHARS_PATTERN = re.compile(r"[\u3000\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff]")


def clean_kami_text(text: Any) -> Any:
    """
    Excel内のテキストをクリーニングする。
    - 全角スペースを半角スペースに正規化
    - 漢字・ひらがな・カタカナの間に挟まった不自然な空白(1-3個)を削除 (例: "氏  名" -> "氏名")
    """
    if not isinstance(text, str):
        return text

    if not text:
        return ""

    # ⚡ Performance: Fast-path for non-Japanese text to avoid regex overhead
    if text.isascii() or not _JAPANESE_CHARS_PATTERN.search(text):
        return text.strip()

    # 全角スペースを半角に
    text = text.replace("\u3000", " ")

    # 漢字・ひらがな・カタカナの間に挟まった1〜3つの空白を削除
    res = _CLEAN_KAMI_TEXT_RE.sub(r"\1", text)

    return res.strip()


def secure_filename(filename: str) -> str:
    """
    Sanitize a string to be used as a filename.
    Keeps only alphanumeric characters, underscores, dashes, and dots.
    Converts spaces to underscores and removes path separators.
    """
    if not filename:
        return "unnamed"

    filename = unicodedata.normalize("NFKD", filename)
    filename = filename.replace(" ", "_")
    filename = _FILENAME_SANITIZE_RE.sub("_", filename)
    filename = filename.strip("._")
    filename = _FILENAME_MULTIPLE_UNDERSCORES_RE.sub("_", filename)
    filename = _FILENAME_MULTIPLE_DOTS_RE.sub(".", filename)

    if not filename or filename in (".", ".."):
        return "unnamed"

    return filename


class CacheManager:
    """SQLiteを使用したキャッシュ永続化マネージャー"""

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._lock = threading.RLock()
        self._conn = None
        self._batch_mode = False
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        if self._conn is None:
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            self._conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
            self._conn.execute("PRAGMA journal_mode=WAL")
            self._conn.execute("PRAGMA synchronous=NORMAL")
        return self._conn

    @contextmanager
    def batch(self) -> Generator["CacheManager", None, None]:
        """複数の操作を1つのトランザクションにまとめる。"""
        with self._lock:
            if self._batch_mode:
                yield self
                return

            self._batch_mode = True
            try:
                yield self
                if self._conn:
                    self._conn.commit()
            except BaseException:  # KeyboardInterruptなどのシグナル時も安全にロールバック
                if self._conn:
                    self._conn.rollback()
                raise
            finally:
                self._batch_mode = False

    def _init_db(self):
        """データベースとテーブルの初期化"""
        with self._lock:
            conn = self._get_conn()
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
            conn.execute("""
                CREATE TABLE IF NOT EXISTS raw_extraction_cache (
                    key TEXT PRIMARY KEY,
                    content TEXT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()

    def get_file_hash(self, file_path: Path) -> str:
        """ファイルの内容からSHA-256ハッシュを生成する"""
        sha256 = hashlib.sha256()
        with open(file_path, "rb") as f:
            while chunk := f.read(8192):
                sha256.update(chunk)
        return sha256.hexdigest()

    async def aget_file_hash(self, file_path: Path) -> str:
        """ファイルの内容からSHA-256ハッシュを非同期で生成する"""
        return await asyncio.to_thread(self.get_file_hash, file_path)

    def get_raw_extraction(self, file_hash: str, include_logic: bool) -> Optional[str]:
        """Excelの生解析結果（HTML/セル情報）をキャッシュから取得"""
        key = f"{file_hash}:logic={include_logic}"
        with self._lock:
            cur = self._get_conn().execute("SELECT content FROM raw_extraction_cache WHERE key = ?", (key,))
            row = cur.fetchone()
            return row[0] if row else None

    async def aget_raw_extraction(self, file_hash: str, include_logic: bool) -> Optional[str]:
        """Excelの生解析結果を非同期で取得"""
        return await asyncio.to_thread(self.get_raw_extraction, file_hash, include_logic)

    def set_raw_extraction(self, file_hash: str, include_logic: bool, content: str):
        """Excelの生解析結果（HTML/セル情報）をキャッシュに保存"""
        key = f"{file_hash}:logic={include_logic}"
        with self._lock:
            conn = self._get_conn()
            conn.execute("INSERT OR REPLACE INTO raw_extraction_cache (key, content) VALUES (?, ?)", (key, content))
            if not self._batch_mode:
                conn.commit()

    async def aset_raw_extraction(self, file_hash: str, include_logic: bool, content: str):
        """Excelの生解析結果を非同期で保存"""
        return await asyncio.to_thread(self.set_raw_extraction, file_hash, include_logic, content)

    def get_vlm_result(self, model: str, prompt: str, image_hash: str) -> Optional[str]:
        """VLMの解析結果をキャッシュから取得"""
        key = f"{model}:{prompt}:{image_hash}"
        with self._lock:
            cur = self._get_conn().execute("SELECT content FROM vlm_cache WHERE key = ?", (key,))
            row = cur.fetchone()
            return row[0] if row else None

    async def aget_vlm_result(self, model: str, prompt: str, image_hash: str) -> Optional[str]:
        """VLMの解析結果を非同期で取得"""
        return await asyncio.to_thread(self.get_vlm_result, model, prompt, image_hash)

    def set_vlm_result(self, model: str, prompt: str, image_hash: str, content: str):
        """VLMの解析結果をキャッシュに保存"""
        key = f"{model}:{prompt}:{image_hash}"
        with self._lock:
            conn = self._get_conn()
            conn.execute("INSERT OR REPLACE INTO vlm_cache (key, content) VALUES (?, ?)", (key, content))
            if not self._batch_mode:
                conn.commit()

    async def aset_vlm_result(self, model: str, prompt: str, image_hash: str, content: str):
        """VLMの解析結果を非同期で保存"""
        return await asyncio.to_thread(self.set_vlm_result, model, prompt, image_hash, content)

    def get_image_data_url(self, image_hash: str) -> Optional[str]:
        """Base64データURLをキャッシュから取得"""
        with self._lock:
            cur = self._get_conn().execute("SELECT data_url FROM image_cache WHERE hash = ?", (image_hash,))
            row = cur.fetchone()
            return row[0] if row else None

    async def aget_image_data_url(self, image_hash: str) -> Optional[str]:
        """Base64データURLを非同期で取得"""
        return await asyncio.to_thread(self.get_image_data_url, image_hash)

    def set_image_data_url(self, image_hash: str, data_url: str):
        """Base64データURLをキャッシュに保存"""
        with self._lock:
            conn = self._get_conn()
            conn.execute("INSERT OR REPLACE INTO image_cache (hash, data_url) VALUES (?, ?)", (image_hash, data_url))
            if not self._batch_mode:
                conn.commit()

    async def aset_image_data_url(self, image_hash: str, data_url: str):
        """Base64データURLを非同期で保存"""
        return await asyncio.to_thread(self.set_image_data_url, image_hash, data_url)

    def get_llm_result(self, model: str, prompt: str, input_text: str) -> Optional[str]:
        """LLMの解析結果をキャッシュから取得"""
        input_hash = hashlib.sha256(input_text.encode("utf-8")).hexdigest()
        key = f"{model}:{hashlib.sha256(prompt.encode('utf-8')).hexdigest()}:{input_hash}"
        with self._lock:
            cur = self._get_conn().execute("SELECT content FROM llm_cache WHERE key = ?", (key,))
            row = cur.fetchone()
            return row[0] if row else None

    async def aget_llm_result(self, model: str, prompt: str, input_text: str) -> Optional[str]:
        """LLMの解析結果を非同期で取得"""
        return await asyncio.to_thread(self.get_llm_result, model, prompt, input_text)

    def set_llm_result(self, model: str, prompt: str, input_text: str, content: str):
        """LLMの解析結果をキャッシュに保存"""
        input_hash = hashlib.sha256(input_text.encode("utf-8")).hexdigest()
        key = f"{model}:{hashlib.sha256(prompt.encode('utf-8')).hexdigest()}:{input_hash}"
        with self._lock:
            conn = self._get_conn()
            conn.execute("INSERT OR REPLACE INTO llm_cache (key, content) VALUES (?, ?)", (key, content))
            if not self._batch_mode:
                conn.commit()

    async def aset_llm_result(self, model: str, prompt: str, input_text: str, content: str):
        """LLMの解析結果を非同期で保存"""
        return await asyncio.to_thread(self.set_llm_result, model, prompt, input_text, content)

    def clear(self):
        """すべてのキャッシュを削除する"""
        with self._lock:
            conn = self._get_conn()
            conn.execute("DELETE FROM vlm_cache")
            conn.execute("DELETE FROM image_cache")
            conn.execute("DELETE FROM llm_cache")
            conn.execute("DELETE FROM raw_extraction_cache")
            conn.commit()
