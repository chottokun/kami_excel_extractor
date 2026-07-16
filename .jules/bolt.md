## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-20 - [clean_kami_text] Expensive Regex for Non-Japanese Text
**Learning:** The `clean_kami_text` utility was executing complex regex and full-width space replacements for every cell, even for pure numeric or ASCII data (common in technical sheets).
**Action:** Use `text.isascii()` as a fast-path to skip Japanese-specific processing for ASCII-only strings, yielding a ~3x speedup for those cases.
