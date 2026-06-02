## 2026-06-02 - [Style Caching in MetadataExtractor]
**Learning:** Caching 'openpyxl' style-dependent computations (CSS strings, unit info, borders) using 'cell.style_id' significantly reduces overhead in large workbooks, as many cells share identical styles. Pre-calculating these once per cell in the main loop further eliminates redundant calls to 'clean_kami_text' and style methods.
**Action:** Always use 'style_id' for style-based caching when working with 'openpyxl' to avoid O(N*S) where N is cells and S is style complexity.
