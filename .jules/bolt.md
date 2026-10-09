## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Single-Pass Bounding Box and Pre-fetched Cell Style Lookups
**Learning:** Combining value and formula sheet row passes using `itertools.zip_longest` and scanning cells right-to-left in `_get_bounding_box` avoids scanning empty cells on wide sheets. Pre-fetching `style_id`, `border_info`, and `coord` per cell eliminates duplicate openpyxl property accesses.
**Action:** Always scan rows right-to-left when searching for maximum bounds, and pass pre-fetched style/coordinate context into cell formatter methods.
