## 2025-05-15 - [Style ID Caching in Excel Extraction]
**Learning:** Parsing cell styles (borders, colors, fonts, number formats) in `openpyxl` is expensive because it involves traversing nested objects. Since many cells in a typical spreadsheet share the same format, `openpyxl` assigns them a common `style_id`. Caching the results of style-to-CSS conversion and unit extraction by this `style_id` dramatically reduces redundant computation.
**Action:** When iterating over large Excel sheets with `openpyxl`, always use `style_id` to cache expensive visual or format-related metadata extraction.
