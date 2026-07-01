## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-20 - [MetadataExtractor] Micro-overhead of openpyxl Property Access
**Learning:** Even seemingly cheap properties like `cell.style_id` and `cell.coordinate` in `openpyxl` incur measurable overhead when accessed multiple times per cell in a large loop.
**Action:** Fetch these properties exactly once per cell at the beginning of the iteration and pass them as arguments to all helper methods that need them.
