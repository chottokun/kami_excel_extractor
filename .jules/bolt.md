## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2026-07-02 - [MetadataExtractor] Single-Pass Right-to-Left Bounding Box Calculation
**Learning:** Scanning Excel rows left-to-right across separate value and formula sheets iterates over every cell needlessly. Scanning row tuples right-to-left in a single combined pass allows breaking immediately upon finding the rightmost non-empty cell in each row.
**Action:** Combine value and formula sheet iterations with `itertools.zip_longest` and iterate right-to-left in `_get_bounding_box`.
