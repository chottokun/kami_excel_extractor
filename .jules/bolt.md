## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Redundant function calls from None default value checks
**Learning:** Passing a pre-calculated `None` value to downstream functions can cause silent re-calculation if the downstream function uses `if param is None:` to check if a value was provided.
**Action:** Use a custom module-level sentinel object (e.g., `_UNSPECIFIED = object()`) as the default parameter value. This allows distinguishing between an unprovided argument and an argument explicitly pre-calculated as `None`.
