## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Caching and SENTINEL-based Loop Optimizations
**Learning:** Even with style caches, repeatedly calling helpers like `_get_unit_info` on the same sheet with `None` values or doing redundant `html.escape` and string cleaning overhead limits performance. Using a custom module-level sentinel `_UNSPECIFIED` for the cell parameter avoids duplicate helper checks, bypassing `clean_kami_text` for non-string types saves function invocation overhead, pre-fetching `style_id` and `coord` avoids openpyxl descriptor property lookup penalties, and caching HTML-escaped cell values in a local `_html_escape_cache` provides a significant cumulative ~11% speedup on large datasets.
**Action:** Pre-fetch descriptor properties, skip string-cleaning logic for numerical/date types, and use a local HTML escape cache within inner generation loops.
