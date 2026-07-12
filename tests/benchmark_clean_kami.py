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
    ascii_text = "This is a simple ASCII string with some spaces " * 10
    japanese_text = "これは 日本語の テキスト です。 氏  名" * 10

    iterations = 100000

    print(f"Benchmarking with {iterations} iterations...")

    # ASCII case
    start = time.time()
    for _ in range(iterations):
        clean_kami_text_original(ascii_text)
    duration_orig_ascii = time.time() - start

    start = time.time()
    for _ in range(iterations):
        clean_kami_text_optimized(ascii_text)
    duration_opt_ascii = time.time() - start

    print(f"ASCII - Original: {duration_orig_ascii:.4f}s, Optimized: {duration_opt_ascii:.4f}s")
    print(f"ASCII Speedup: {duration_orig_ascii / duration_opt_ascii:.2f}x")

    # Japanese case
    start = time.time()
    for _ in range(iterations):
        clean_kami_text_original(japanese_text)
    duration_orig_ja = time.time() - start

    start = time.time()
    for _ in range(iterations):
        clean_kami_text_optimized(japanese_text)
    duration_opt_ja = time.time() - start

    print(f"Japanese - Original: {duration_orig_ja:.4f}s, Optimized: {duration_opt_ja:.4f}s")
    print(f"Japanese overhead: {duration_opt_ja / duration_orig_ja:.2f}x")

if __name__ == "__main__":
    benchmark()
