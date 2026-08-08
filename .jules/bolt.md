## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Coordinate Pre-fetching, Sentinel Unit parameter, and HTML Escape Cache
**Learning:** Accessing openpyxl cell coordinate dynamically (via `cell.coordinate`) is computed on-the-fly and extremely slow. Furthermore, repeating `html.escape` and string replacement operations in HTML TD tag generation is highly redundant.
**Action:** Pre-calculate and pass the coordinate and resolved unit parameters explicitly, and cache the HTML escaping of cleaned values inside a `self._html_escape_cache` lookup table. Use a custom sentinel `_UNSPECIFIED` for the default unit argument to handle None values cleanly without triggering recalculations.
