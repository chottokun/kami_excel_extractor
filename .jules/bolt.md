## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-07-14 - [clean_kami_text] ASCII Fast-Path Optimization
**Learning:** The `clean_kami_text` function, which uses regex to normalize Japanese text, is a bottleneck in the cell extraction loop. However, for pure ASCII strings (common in data-heavy sheets), these regex operations are redundant and expensive.
**Action:** Use `str.isascii()` as a fast-path to bypass regex and full-width space replacement. This yielded a ~31x speedup for ASCII text with negligible overhead for Japanese text.
