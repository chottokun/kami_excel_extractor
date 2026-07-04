## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-07-04 - [MetadataExtractor & Utils] Cell Loop and Text Cleaning Optimization
**Learning:** Redundant `getattr(cell, "style_id")` and `cell.coordinate` access in tight loops, along with unnecessary regex operations on ASCII text in `clean_kami_text`, significantly impact sheet processing speed.
**Action:** Pre-calculate `style_id` and `coord` once per cell. Implement an `isascii()` fast-path in `clean_kami_text` to bypass expensive CJK cleaning for simple text. Skip calling string cleaners for date/datetime objects.
