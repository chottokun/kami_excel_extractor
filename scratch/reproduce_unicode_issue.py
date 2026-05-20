import subprocess


def test_unicode_decode_error_simulation():
    # External tool that outputs invalid UTF-8 (e.g., using latin-1 encoding for a byte that's not valid UTF-8)
    # The byte 0xFF is not a valid start of a UTF-8 sequence.
    try:
        # This will fail if errors='replace' is not used and text=True
        print("Running WITHOUT errors='replace'...")
        res = subprocess.run(
            ["python3", "-c", "import sys; sys.stderr.buffer.write(b'Invalid: \\xff\\n')"],
            capture_output=True,
            text=True,
        )
        print(f"STDOUT: {res.stdout}")
        print(f"STDERR: {res.stderr}")
    except UnicodeDecodeError as e:
        print(f"Caught expected UnicodeDecodeError: {e}")

    try:
        print("\nRunning WITH errors='replace'...")
        res = subprocess.run(
            ["python3", "-c", "import sys; sys.stderr.buffer.write(b'Invalid: \\xff\\n')"],
            capture_output=True,
            text=True,
            errors="replace",
        )
        print(f"STDOUT: {res.stdout}")
        print(f"STDERR: {res.stderr}")
    except UnicodeDecodeError as e:
        print(f"Caught UNEXPECTED UnicodeDecodeError: {e}")


if __name__ == "__main__":
    test_unicode_decode_error_simulation()
