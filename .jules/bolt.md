## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-20 - [clean_kami_text] Redundant Regex on ASCII Strings
**Learning:** The `clean_kami_text` function performs Japanese-specific regex cleaning. Running this on pure ASCII strings (common in many Excel sheets) is a significant waste of CPU.
**Action:** Add an `isascii()` fast path to skip regex and string replacements for ASCII-only content.
