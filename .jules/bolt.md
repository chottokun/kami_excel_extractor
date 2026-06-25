## 2025-05-15 - [Style Caching in MetadataExtractor]
**Learning:** Accessing cell style attributes (font, borders, fill) in `openpyxl` for every cell is a significant bottleneck during large sheet extraction because it involves repeated attribute lookups and object proxies. `cell.style_id` provides a reliable integer key to cache these properties per-workbook.
**Action:** Use a dictionary-based `style_id` cache in Excel extraction loops to avoid redundant CSS string generation and unit inference for identically styled cells.
