## 2025-05-15 - [MetadataExtractor Style Caching]
**Learning:** `openpyxl` cell styles are computationally expensive to parse individually (borders, fill, font). Since most spreadsheets reuse a small set of styles across thousands of cells, caching the parsed CSS and metadata results keyed by `cell.style_id` provides a massive performance boost.
**Action:** Always check for `style_id` or similar internal identifiers when iterating over large datasets with repetitive formatting to avoid redundant property access and string formatting.
