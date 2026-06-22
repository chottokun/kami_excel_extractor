## 2026-06-22 - [Excel Style Caching]
**Learning:** Excel formatting (styles, number formats, borders) is often highly redundant across thousands of cells in a worksheet. Using the internal `style_id` of `openpyxl` cells to cache pre-computed CSS and metadata significantly reduces execution time by avoiding expensive property lookups and repeated string formatting.
**Action:** Always check for property-based redundancy when iterating over large datasets (like spreadsheet cells) and implement a scoped cache (e.g., per-worksheet) to reuse expensive computations.
