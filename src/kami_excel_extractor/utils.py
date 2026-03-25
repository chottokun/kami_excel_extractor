import re
import unicodedata

def secure_filename(filename: str) -> str:
    """
    Sanitize a string to be used as a filename.
    Keeps only alphanumeric characters, underscores, dashes, and dots.
    Converts spaces to underscores and removes path separators.
    """
    if not filename:
        return "unnamed"

    # Normalize unicode characters to NFKD (separate base characters from marks)
    filename = unicodedata.normalize('NFKD', filename)

    # Replace spaces with underscores
    filename = filename.replace(" ", "_")

    # Remove anything that isn't alphanumeric, underscore, dash, or dot
    # We allow Japanese/other characters if they are alphanumeric according to unicode
    # But for maximum security, we might want to restrict to ASCII if we don't care about Japanese names.
    # However, this tool works with Japanese Excel files, so we should support Japanese.

    # Keep alphanumeric (including Japanese), underscores, dashes, and dots.
    # [^\w\.\-] where \w includes Unicode word characters.
    filename = re.sub(r'[^\w\.\-]', '_', filename)

    # Strip leading/trailing dots and underscores
    filename = filename.strip("._")

    # Remove multiple consecutive underscores/dots
    filename = re.sub(r'__+', '_', filename)
    filename = re.sub(r'\.\.+', '.', filename)

    # If the filename becomes empty, use a default
    if not filename or filename in (".", ".."):
        return "unnamed"

    return filename
