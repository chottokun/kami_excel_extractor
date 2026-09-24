## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] `_UNSPECIFIED` Sentinel and Reverse Row Bounding Box Scanning
**Learning:** Defaulting parameter `unit=None` in `_cell_to_html_td` caused fallback calls to `_get_unit_info` for every cell without a unit format (~125,000 extra calls). Using a sentinel object (`_UNSPECIFIED`) avoids duplicate function calls for `None` values. Additionally, scanning row cells in reverse (right-to-left) allows breaking out of column iteration as soon as the rightmost data/formula cell is found.
**Action:** Always use unique sentinel objects for optional parameters where `None` is a valid pre-computed value to avoid unintended fallback computations, and scan bounding box rows in reverse to exit early.
