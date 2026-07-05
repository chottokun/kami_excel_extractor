## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] ASCII Fast-path and Reduced Attribute Access
**Learning:** `clean_kami_text` is called for every cell, and its regex-based Japanese cleaning is expensive for non-Japanese (ASCII) text. Also, passing `style_id` to helper methods avoids redundant `getattr` calls in tight loops.
**Action:** Use `text.isascii()` as a fast-path in text cleaning. Pre-calculate `style_id` and `coordinate` in loops to minimize `openpyxl` object property access overhead.
