## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-10-24 - [MetadataExtractor] O(R*C) Full Grid Bounding Box Scan Bottleneck
**Learning:** For large, sparse sheets, using `ws.iter_rows()` to scan the entire grid to determine sheet bounds causes massive overhead, especially when many cells are empty. Directly iterating over `ws._cells` dictionary elements is O(N) (where N is only populated cells), resulting in a major performance boost (>4x faster bounds checking).
**Action:** Use internal structure `ws._cells` with a fallback to `ws.iter_rows()` for O(N) bounding box detection.
