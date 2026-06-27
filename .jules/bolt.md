## 2025-05-15 - style_id Caching in openpyxl
**Learning:** `openpyxl` cells share a `style_id` if they have identical formatting. Accessing cell style attributes (like `font`, `border`, `fill`) triggers internal proxy object creation and lookups which are expensive when repeated for thousands of cells.
**Action:** Cache pre-calculated style information (CSS strings, border maps, unit types) using `cell.style_id` as the key. This can reduce extraction logic time by over 75% for typical sheets with repetitive styling. Ensure the cache is reset per workbook to avoid ID collisions.
