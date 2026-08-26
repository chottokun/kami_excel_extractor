## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-08-26 - [MetadataExtractor] Default `None` Arguments Trigger Fallback Lookups
**Learning:** Using `None` as a default parameter value in helper functions like `_cell_to_html_td` causes explicit `None` arguments (e.g. `unit=None` passed for default unformatted cells) to evaluate `if unit is None:` as True, doubling function calls and property lookups across cell loops.
**Action:** Use a custom sentinel `_UNSPECIFIED = object()` for optional parameters to distinguish between unpassed parameters and explicitly passed `None` values.
