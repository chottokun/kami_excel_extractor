## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Redundant openpyxl Coordinate Lookups and Repetitive HTML Escaping
**Learning:** Accessing openpyxl's `cell.coordinate` is surprisingly slow because it dynamically formats column letters on every access. Furthermore, executing `html.escape` and string replacement for every cell value is extremely repetitive and slow, especially on large sheets with many empty or identical values.
**Action:** Pre-calculate `coord = cell.coordinate` exactly once inside the iteration loop and pass it downstream. Use an instance-bound `self._html_escape_cache` dictionary to cache and instantly retrieve escaped HTML cell values, yielding up to an 8.7x micro-benchmark speedup.
