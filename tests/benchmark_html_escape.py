import timeit
import html

# Original logic
def original_logic(val_str):
    if not val_str:
        return ""
    elif any(c in val_str for c in '&<>\n"\''):
        return html.escape(val_str).replace('\n', '<br>')
    else:
        return val_str

# Optimized logic (directly escaping)
def optimized_logic(val_str):
    if not val_str:
        return ""
    return html.escape(val_str).replace('\n', '<br>')

# Test inputs
inputs = {
    "empty": "",
    "normal_short": "Hello World",
    "normal_long": "This is a much longer string that doesn't have any special characters but is still quite long to process." * 10,
    "special_short": "Hello <World>",
    "special_long": "This is a much longer string that DOES have some & special characters <here> and there, and also some\nnewlines." * 10
}

def run_benchmark(logic_func, name):
    print(f"--- Benchmarking: {name} ---")
    for key, val in inputs.items():
        t = timeit.timeit(lambda: logic_func(val), number=100000)
        print(f"{key:15}: {t:.6f}s")
    print()

if __name__ == "__main__":
    run_benchmark(original_logic, "Original Logic (any())")
    run_benchmark(optimized_logic, "Optimized Logic (direct escape)")
