
## 2026-06-23 - [MetadataExtractor style caching]
**Learning:** In `openpyxl`, cell styles are indexed by `style_id`. Many cells share the same style. Caching the results of `_get_cell_style_string`, `_get_unit_info`, and border extraction by `style_id` avoids redundant processing for thousands of cells.
**Action:** Use `cell.style_id` as a cache key when extracting visual metadata from Excel workbooks to achieve significant performance gains (measured ~35% speedup).
