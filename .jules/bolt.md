## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [clean_kami_text] Inefficient String and Regex Handling for Pure ASCII Content
**Learning:** Calling `.replace()` and regex `.sub()` operations on strings that contain only ASCII characters (e.g., coordinates, numbers, identifiers, empty strings) creates significant overhead because they never match Japanese character classes or full-width spaces.
**Action:** Add a fast-path checking `text.isascii()` to immediately strip and return ASCII strings, bypassing the expensive regex engine entirely. This yields a ~3-80x micro-benchmark speedup.
