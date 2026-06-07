## 2025-05-15 - [Style Caching in MetadataExtractor]
**Learning:** In openpyxl, many cells share the same style, identifiable by `cell.style_id`. Recalculating CSS strings, units, and borders for every cell is a major bottleneck during extraction.
**Action:** Use a per-extraction cache keyed by `cell.style_id` to store pre-computed style attributes (CSS, unit, borders, bold flag). This achieved a ~36-44% speedup in cell processing and ~24% overall extraction speedup on realistic files.
