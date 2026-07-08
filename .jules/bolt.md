## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-07-08 - [Extraction Engine] Loop and Utility Optimization
**Learning:** Redundant property access (e.g., cell.style_id, cell.coordinate) in large sheet loops and expensive regex calls for ASCII text in clean_kami_text are significant bottlenecks.
**Action:** Pre-calculate cell attributes once per loop iteration, pass them to helpers, and implement isascii() fast-paths for text processing.
