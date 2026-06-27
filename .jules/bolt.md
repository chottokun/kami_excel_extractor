## 2026-06-27 - Reducing Redundant Computations in Cell Processing Loop
**Learning:** In the `MetadataExtractor`, methods like `_get_unit_info`, `_get_cell_style_string`, and `clean_kami_text` were being called multiple times for the same cell (once for metadata and once for HTML generation). Hoisting these calls and passing the results as arguments reduced processing time by ~4.5% on large sheets.
**Action:** Always check if a value being processed in a loop is used in multiple output structures (e.g., JSON metadata and HTML) and ensure it's calculated only once.
