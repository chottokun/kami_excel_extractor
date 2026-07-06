## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-15 - [MetadataExtractor] ASCII Fast-path and Font Property Caching
**Learning:** String cleaning with regex and accessing `openpyxl` font properties (`cell.font.i`) are significant bottlenecks during large sheet extraction. ASCII-only strings can safely skip Japanese-specific cleaning.
**Action:** Implement `isascii()` fast-path in `clean_kami_text` and cache font properties by `style_id` to minimize expensive object attribute lookups.
