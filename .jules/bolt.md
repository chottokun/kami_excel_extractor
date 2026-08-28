## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Early Break on Reverse Row Scanning in Bounding Box Calculation
**Learning:** In openpyxl worksheet row iterations (`iter_rows`), scanning cells left-to-right checks all cells up to max column, whereas scanning cells right-to-left allows breaking immediately upon finding the first non-None data cell per row. Additionally, combining values and formula sheet iterations into a single pass avoids redundant full sheet traversals.
**Action:** In bounding box calculations, zip values and formulas generators into a single pass and scan each row in reverse (`range(len-1, -1, -1)`) to exit row iteration as early as possible.
