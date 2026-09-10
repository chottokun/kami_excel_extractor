## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Reverse Scan for Bounding Box Rightmost Column
**Learning:** In Excel sheets with sparse data, scanning columns left-to-right checks many empty cells. Scanning each row right-to-left finds the rightmost non-None cell immediately and breaks out early.
**Action:** Use reverse index scanning (`range(len(row) - 1, -1, -1)`) and single-pass `itertools.zip_longest` when determining worksheet bounding boxes.
