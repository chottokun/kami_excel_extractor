## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [HTML Generation] Redundant Method Invocations for Default Parameter Values
**Learning:** In cell iteration, using `None` as a default parameter value in helper functions can trigger redundant computations (like `_get_unit_info`) if `None` is also a valid returned value passed from parent scopes.
**Action:** Use a custom sentinel object (`_UNSPECIFIED = object()`) as the default parameter value instead of `None` so that explicit `None` values do not trigger re-computation fallback logic.
