import time
import re
from typing import Any

_CLEAN_KAMI_TEXT_RE = re.compile(
    r"([\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])\s{1,3}(?=[\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])"
)

def clean_kami_text_original(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    text = text.replace("\u3000", " ")
    res = _CLEAN_KAMI_TEXT_RE.sub(r"\1", text)
    return res.strip()

def clean_kami_text_optimized(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    if text.isascii():
        return text.strip()
    text = text.replace("\u3000", " ")
    res = _CLEAN_KAMI_TEXT_RE.sub(r"\1", text)
    return res.strip()

def benchmark():
    ascii_text = "This is a test string with some spaces " * 10
    japanese_text = "これは　テスト　です。氏　名" * 10

    iterations = 100000

    print(f"Benchmarking with {iterations} iterations...")

    # ASCII Original
    start = time.time()
    for _ in range(iterations):
        clean_kami_text_original(ascii_text)
    duration_ascii_orig = time.time() - start
    print(f"ASCII Original: {duration_ascii_orig:.4f}s")

    # ASCII Optimized
    start = time.time()
    for _ in range(iterations):
        clean_kami_text_optimized(ascii_text)
    duration_ascii_opt = time.time() - start
    print(f"ASCII Optimized: {duration_ascii_opt:.4f}s")
    print(f"ASCII Improvement: {duration_ascii_orig / duration_ascii_opt:.2f}x faster")

    # Japanese Original
    start = time.time()
    for _ in range(iterations):
        clean_kami_text_original(japanese_text)
    duration_jp_orig = time.time() - start
    print(f"Japanese Original: {duration_jp_orig:.4f}s")

    # Japanese Optimized
    start = time.time()
    for _ in range(iterations):
        clean_kami_text_optimized(japanese_text)
    duration_jp_opt = time.time() - start
    print(f"Japanese Optimized: {duration_jp_opt:.4f}s")
    print(f"Japanese Overhead: {duration_jp_opt / duration_jp_orig:.4f}x (lower is better, 1.0 means no overhead)")

if __name__ == "__main__":
    benchmark()
