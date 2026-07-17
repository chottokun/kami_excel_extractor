## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] High-Performance and Robust Worksheet Bounding Box Scanning
**Learning:** Sequential scans of worksheets via `ws.iter_rows(values_only=True)` process every empty cell in the bounding box grid, causing significant overhead in sparse sheets or large layouts. Directly iterating over `ws._cells` avoids empty grid cell generation and achieves up to ~6x speedups. However, to guarantee robustness against future `openpyxl` updates and `read_only` modes, a safe fallback to `iter_rows` should always be included.
**Action:** Always check for `hasattr(ws, "_cells") and isinstance(ws._cells, dict)` before using the fast-path direct dictionary iteration, falling back gracefully to the standard public API.
