## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Sentinel Object for Default Parameter Fallback Bypass
**Learning:** When optional parameters take `None` as a valid pre-computed result (e.g. `unit=None` when a cell has no unit), defaulting parameters to `None` causes helper methods to re-trigger fallback calculation logic. In large loops, this doubled helper calls.
**Action:** Use a unique sentinel object (`_UNSPECIFIED = object()`) as the default argument value so pre-computed `None` values do not trigger re-computation.
