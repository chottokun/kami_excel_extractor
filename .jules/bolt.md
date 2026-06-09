## 2026-06-09 - [Style ID Caching in MetadataExtractor]
**Learning:** `openpyxl` cells share `style_id` for identical formatting. Calculating CSS strings, units, and borders for every cell is redundant and slow for large sheets.
**Action:** Implement a per-sheet cache keyed by `cell.style_id` to store pre-calculated style attributes, significantly reducing extraction time.
