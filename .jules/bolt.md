## 2025-06-10 - [Style Caching in MetadataExtractor]
**Learning:** In Excel spreadsheets, many cells share the same style (font, borders, fill, number format). Calculating these visual attributes for every single cell is extremely redundant and becomes a major bottleneck for large sheets. By caching style strings and metadata keyed by `cell.style_id`, we can reduce O(N*M) calculations to O(UniqueStyles).
**Action:** Always check if row/cell properties in openpyxl can be cached using `style_id` when performing batch processing of spreadsheets.
