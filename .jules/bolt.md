## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] HTML Escape Caching and Sentinel Bystanders
**Learning:** Repeatedly calling `html.escape` on same/similar string values (e.g. empty strings, common words, formulas) and evaluating `_get_unit_info` on pre-calculated `None` values is a significant cpu and function call bottleneck in cell-heavy loops.
**Action:** Use a fast HTML escape cache keyed by raw strings, and use a unique sentinel object `_UNSPECIFIED` for parameters with valid `None` defaults so that pre-calculated `None` values are not computed twice.
