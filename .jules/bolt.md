## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-07-01 - [MetadataExtractor] Over-Escaping and Double-Calculation Bottlenecks in TD Formatting
**Learning:** Even with pre-calculated cells, passing `unit=None` triggered fallback lookups due to standard default checks. Additionally, coordinate lookup `.coordinate` and standard `html.escape` string operations on high-frequency cells (like empty cells) incur heavy overhead.
**Action:** Implement `_UNSPECIFIED` sentinel to bypass duplicate lookups, pre-fetch and pass coordinates, and use a simple local cache (`self._html_escape_cache`) for escaping.
