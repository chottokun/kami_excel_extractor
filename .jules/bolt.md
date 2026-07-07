## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Loop Optimization and ASCII Fast-path
**Learning:** The cell processing loop in `MetadataExtractor` was performing redundant lookups and function calls (e.g., `getattr(cell, "style_id")`, `cell.coordinate`) for every cell. Additionally, `clean_kami_text` was applying expensive regex to all strings, including plain ASCII ones.
**Action:** Pre-calculate `style_id` and `coord` once per cell. Implement an `isascii()` fast-path in text cleaning. These changes reduced benchmark duration by ~18.5%.
