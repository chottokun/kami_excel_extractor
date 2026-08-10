## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-03-05 - [clean_kami_text] Bypass Expensive Regex and Replace Operations for ASCII Strings
**Learning:** `clean_kami_text` is heavily called (over 120k times for a 500x50 sheet benchmark) and originally performs full-width space replacement and regex matching even for ASCII-only strings like cells containing "Data 1-1", which is a major CPU bottleneck. Since Japanese regex and full-width spaces are non-ASCII, pure ASCII strings can never trigger these replacements.
**Action:** Add an `isascii()` fast-path to short-circuit the execution for ASCII-only content, which provides a massive 4.5x - 58.x speedup for ASCII data while preserving exact behavioral correctness.
