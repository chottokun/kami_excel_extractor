import timeit
from typing import Any
import re

_CLEAN_KAMI_TEXT_RE = re.compile(
    r"([\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])\s{1,3}(?=[\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])"
)

def original_clean(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    text = text.replace("\u3000", " ")
    res = _CLEAN_KAMI_TEXT_RE.sub(r"\1", text)
    return res.strip()

def optimized_clean(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    if text.isascii():
        return text.strip()
    text = text.replace("\u3000", " ")
    res = _CLEAN_KAMI_TEXT_RE.sub(r"\1", text)
    return res.strip()

inputs = {
    "empty": "",
    "ascii_short": "Hello World",
    "ascii_long": "This is a much longer string that doesn't have any non-ascii characters but is still quite long to process." * 10,
    "ascii_number": "12345.67",
    "japanese_short": "氏  名",
    "japanese_long": "これは日本語のテスト用の長い文字列です。漢字、ひらがな、カタカナが混ざっています。" * 10,
}

def run_benchmark(clean_func, name):
    print(f"--- Benchmarking: {name} ---")
    for key, val in inputs.items():
        t = timeit.timeit(lambda: clean_func(val), number=100000)
        print(f"{key:15}: {t:.6f}s")
    print()

if __name__ == "__main__":
    run_benchmark(original_clean, "Original clean_kami_text")
    run_benchmark(optimized_clean, "Optimized clean_kami_text with isascii()")
