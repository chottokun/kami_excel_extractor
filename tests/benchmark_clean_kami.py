import time
import timeit
from kami_excel_extractor.utils import clean_kami_text

def benchmark():
    inputs = [
        "氏  名",
        "氏   名",
        "あ　い",
        "  trimmed  ",
        "123",
        "Hello World",
        "A   B   C",
        "1234567890",
        "This is a longer pure ASCII text string with some numbers 42 and punctuation! It is extremely common in Excel sheets.",
        "This is mixed text: 氏  名, with some Japanese characters and formatting."
    ]

    print("Running clean_kami_text micro-benchmark...")
    start_time = time.time()

    # Run 100,000 times
    for _ in range(100000):
        for val in inputs:
            clean_kami_text(val)

    end_time = time.time()
    duration = end_time - start_time
    print(f"Micro-benchmark duration: {duration:.4f} seconds")

if __name__ == "__main__":
    benchmark()
