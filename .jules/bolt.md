## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Coordinate Pre-fetching, HTML Escape Caching, and Sentinel Units
**Learning:** Pre-fetching dynamic coordinate property access once, caching HTML-escaped cell values (avoiding repeated string escaping overhead for identical strings), and using a custom sentinel `_UNSPECIFIED` to distinguish missing units from explicitly `None` units avoids redundant calls and property access inside the main cell processing loops.
**Action:** Always pre-fetch cell properties (e.g., coordinate, style_id) once inside hot loops, cache HTML-escaped representations of common repeating values, and use sentinels to prevent default fallbacks from overriding pre-calculated results.
