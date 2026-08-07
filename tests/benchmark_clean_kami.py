import timeit
import re
from typing import Any

_CLEAN_KAMI_TEXT_RE = re.compile(
    r"([\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])\s{1,3}(?=[\u4e00-\u9faf\u3040-\u309f\u30a0-\u30ff])"
)

def original_clean_kami_text(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    text = text.replace("\u3000", " ")
    res = _CLEAN_KAMI_TEXT_RE.sub(r"\1", text)
    return res.strip()

def optimized_clean_kami_text(text: Any) -> Any:
    if not isinstance(text, str):
        return text
    if text.isascii():
        return text.strip()
    text = text.replace("\u3000", " ")
    res = _CLEAN_KAMI_TEXT_RE.sub(r"\1", text)
    return res.strip()

# Test inputs
inputs = {
    "ascii_short_no_spaces": "Data_1-1",
    "ascii_short_spaces": "  Some ASCII text here  ",
    "ascii_long": "This is a long ASCII string with many spaces and words but no Japanese characters" * 10,
    "japanese_short": "氏  名",
    "japanese_long": "日本語のテキストに　全角スペースや  半角スペースが混ざっているもの" * 5,
}

def run_benchmark():
    print("--- Micro-benchmark: clean_kami_text ---")
    for key, val in inputs.items():
        t_orig = timeit.timeit(lambda: original_clean_kami_text(val), number=100000)
        t_opt = timeit.timeit(lambda: optimized_clean_kami_text(val), number=100000)
        speedup = t_orig / t_opt if t_opt > 0 else 0
        print(f"{key:25} | Original: {t_orig:.6f}s | Optimized: {t_opt:.6f}s | Speedup: {speedup:.2f}x")

if __name__ == "__main__":
    run_benchmark()
