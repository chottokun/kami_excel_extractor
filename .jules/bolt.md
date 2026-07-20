## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-03-05 - [clean_kami_text] Redundant Function Calls and Regex Overhead
**Learning:** Calling `clean_kami_text` for numeric/boolean values is highly redundant. Furthermore, running multi-byte/regex replacements on pure ASCII strings incurs massive unnecessary string and regex processing overhead in the hot cell loop.
**Action:** Add an `.isascii()` check fast-path to `clean_kami_text` to bypass regex entirely for ASCII values, and explicitly type-check cell values in `extractor.py` loops so non-string values bypass `clean_kami_text` completely.
