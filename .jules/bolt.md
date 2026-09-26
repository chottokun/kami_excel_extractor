## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Sentinel Defaults & Right-to-Left Boundary Scanning
**Learning:** In openpyxl worksheet bounds calculations, scanning cells right-to-left allows early loop exit at the first non-empty cell in each row. Furthermore, helper methods using `if arg is None:` fallbacks re-trigger calculations when `None` is a valid pre-computed value; using a sentinel object `_UNSPECIFIED = object()` prevents redundant computations.
**Action:** Use right-to-left scanning for row boundary scans and sentinel defaults for optional pre-calculated parameters in cell loops.
