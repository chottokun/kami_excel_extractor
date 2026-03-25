from kami_excel_extractor.utils import secure_filename

def test_secure_filename():
    test_cases = [
        ("NormalName", "NormalName"),
        ("Spaces In Name", "Spaces_In_Name"),
        ("../../../tmp/evil", "tmp_evil"),
        ("..", "unnamed"),
        (".", "unnamed"),
        ("", "unnamed"),
        ("!@#$%^&*()", "unnamed"),
        ("Sheet (1)", "Sheet_1"),
        ("シート１", "シート1"),  # Japanese support
        ("My.File.xlsx", "My.File.xlsx"),
        ("___multiple___", "multiple"),
        ("...dots...", "dots"),
    ]

    for input_str, expected_output in test_cases:
        output = secure_filename(input_str)
        print(f"Input: '{input_str}' -> Output: '{output}' (Expected: '{expected_output}')")
        assert output == expected_output or (output == "unnamed" and expected_output == "unnamed")

if __name__ == "__main__":
    import sys
    import os
    # Add src to sys.path if needed
    sys.path.append(os.path.join(os.getcwd(), "src"))
    test_secure_filename()
