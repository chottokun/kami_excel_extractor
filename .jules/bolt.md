## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-07-13 - [MetadataExtractor] Optimized Loop and Style Caching
**Learning:** Hoisting constants and pre-calculating cell properties like `style_id` and `coordinate` (which are computed properties in openpyxl) provides a significant speed boost in tight loops. Propagating these values avoids redundant `getattr` and string formatting calls.
**Action:** Always pre-fetch `style_id` and pass it to style-related helpers. Hoist dictionaries used for mapping values to module-level constants.
