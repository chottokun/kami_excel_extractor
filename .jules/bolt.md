## 2026-06-16 - Optimize MetadataExtractor with style caching
**Learning:** openpyxl cell attribute lookups (styles, fonts, borders, number formats) are expensive because they involve traversing complex object graphs for every cell. In many Excel files, a large number of cells share the same style.
**Action:** Use `cell.style_id` as a cache key to store and reuse computed visual attributes (CSS strings, units, border info, font weight) within a single workbook extraction session. This significantly reduces redundant computations in the main iteration loop.
