## 2025-05-15 - [Style Caching in MetadataExtractor]
**Learning:** repeated access to openpyxl cell style properties (font, fill, border) is expensive because they are often proxied or re-instantiated. Since many cells share the same style_id, caching these derived values (CSS strings, units) significantly improves performance.
**Action:** Use cell.style_id as a cache key for cell-level formatting information within a workbook scope.
