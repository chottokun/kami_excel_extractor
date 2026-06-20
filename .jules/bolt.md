## 2025-05-15 - [Style ID Caching in Metadata Extraction]
**Learning:** `openpyxl`'s `cell.style_id` is a highly efficient key for caching style-related calculations (CSS generation, border info, unit extraction, font weight). In large Excel sheets, redundant style processing for every cell is a major bottleneck.
**Action:** Always leverage `style_id` to cache formatting logic in spreadsheet processing engines to reduce complexity from O(cells) to O(unique styles).
