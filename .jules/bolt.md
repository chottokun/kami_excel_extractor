## 2025-05-14 - [Style Caching in openpyxl]
**Learning:** `openpyxl` cells have a `style_id` attribute that represents a unique combination of formatting (font, fill, border, etc.). Computing CSS strings and parsing units for every cell is expensive. Caching these results by `style_id` and resetting the cache per workbook provides a massive performance boost for styled spreadsheets.
**Action:** When extracting metadata or HTML from Excel, use `cell.style_id` as a cache key for any style-dependent calculations.
