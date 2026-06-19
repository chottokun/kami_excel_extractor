## 2025-05-15 - MetadataExtractor style caching
**Learning:** Calculating CSS strings and cell metadata (borders, units) for every cell in a large spreadsheet is a major bottleneck in openpyxl-based extractors. openpyxl's `style_id` property allows for efficient grouping of cells with identical formatting.
**Action:** Use `cell.style_id` to cache style-related computations at the worksheet level. Refactor internal HTML generation methods to accept these pre-calculated values to avoid redundant property access and logic execution.
