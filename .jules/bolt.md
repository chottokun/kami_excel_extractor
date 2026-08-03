## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Redundant Function Calls and Slow openpyxl Coordinate Lookups
**Learning:** Even with caches, redundantly invoking cell properties like `cell.coordinate` inside cell extraction loops causes significant overhead. Furthermore, passing `None` as optional parameter defaults (like `unit`) forces duplicate calculations inside sub-helpers because `is None` triggers recalculated lookups.
**Action:** Pre-fetch and pass coordinate strings, use dictionary caches for HTML escaping of clean cell values, and implement sentinel-based defaults (`_UNSPECIFIED = object()`) to avoid duplicate logic triggers when `None` is a valid pre-computed value.
