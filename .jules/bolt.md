## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Slow Cell coordinate Access and Redundant HTML Escaping
**Learning:** `cell.coordinate` is accessed multiple times per cell in `MetadataExtractor`, triggering slow string construction in `openpyxl`. In addition, repeatedly escaping strings with `html.escape` (especially highly redundant values like spaces, numeric values, etc.) takes up measurable CPU time in large loops.
**Action:** Pre-calculate `cell.coordinate` once per cell and pass it to downstream methods. Cache `html.escape` results for clean cell values in an instance-level cache to skip redundant escaping.
