## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-09-05 - [MetadataExtractor] Early-Exit Row Boundary Scan and HTML Escape Caching
**Learning:** Sequential left-to-right scanning of worksheet cells for bounding box calculation inspects thousands of trailing empty cells unnecessarily, and duplicate scans for formula sheets double this cost. Caching HTML-escaped cell strings prevents thousands of redundant string replacement operations per extraction run.
**Action:** Single-pass `zip_longest` over value and formula worksheets combined with right-to-left row scanning allows early termination at the rightmost data cell. Cache HTML-escaped strings for cell values.
