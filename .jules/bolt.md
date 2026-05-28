## 2025-05-28 - Excel Style Caching
**Learning:** `openpyxl` cells often share identical formatting, which is identified by a `style_id`. Recalculating CSS strings, borders, and unit information for every cell in a large spreadsheet (e.g., 2000x30) is a major bottleneck. Caching these computations by `style_id` reduced extraction time by ~55% (from 4.0s to 1.8s) in this codebase.
**Action:** When extracting data from spreadsheets using `openpyxl`, always leverage `style_id` to cache formatting-related computations. Ensure the cache is reset between workbooks to prevent style leakage.
