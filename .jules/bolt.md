## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-10-06 - [MetadataExtractor] `_UNSPECIFIED` Sentinel and Pre-Fetched Attributes Cut 4.6M Calls
**Learning:** In cell iteration loops, checking `if unit is None` triggered redundant fallback property lookups for cells with no unit format (most cells). Using a sentinel default `_UNSPECIFIED = object()` prevents re-calculating `None` values. Additionally, pre-fetching `style_id`, `border_info`, `coord`, and caching escaped HTML cell strings in `_val_escape_cache` eliminated over 4.6M function calls per run.
**Action:** Always use sentinel objects for default parameters when `None` is a valid pre-computed value, and pre-fetch dynamic cell coordinates and style attributes before inner processing.
