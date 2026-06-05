## 2025-06-05 - Optimize style extraction with caching
**Learning:** `openpyxl`'s style and unit info extraction is expensive when called for every cell. Since Excel workbooks often share styles across many cells (header rows, data columns), caching these results by `cell.style_id` significantly improves performance.
**Action:** Always check for shared attributes like `style_id` when iterating over large grids in `openpyxl` to avoid redundant computation.
