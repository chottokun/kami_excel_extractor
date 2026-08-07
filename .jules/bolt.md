## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [clean_kami_text] Redundant Japanese Space and Regex Normalization for Pure ASCII Strings
**Learning:** For pure ASCII strings (IDs, numbers, english-only strings), running regex search/replace and unicode full-width space substitution in `clean_kami_text` is a redundant and slow operation.
**Action:** Use `text.isascii()` to immediately detect pure ASCII strings and return `text.strip()`, bypassing the expensive regex processing. This provides up to a ~4x to 60x micro-performance speedup.
