## 2025-05-15 - [MetadataExtractor] Redundant Property Access and Computation in Cell Loop
**Learning:** Accessing `openpyxl` cell properties like `cell.font.b` and repeatedly computing styles/units/clean-text for both metadata and HTML generation is a major bottleneck in large sheets.
**Action:** Pre-calculate these values once per cell and pass them to downstream methods. Use caches keyed by `style_id` for expensive property lookups that are shared across cells.

## 2025-05-16 - [MetadataExtractor] Slow Coordinate/Style Lookups and Redundant HTML Escaping
**Learning:** Retrieving `style_id` and `cell.coordinate` properties on openpyxl cells are expensive descriptors that significantly slow down hot loops when accessed repeatedly. In addition, calling `html.escape` repeatedly on duplicate cell contents adds up in string allocation and processing overhead.
**Action:** Pre-fetch `cell.style_id` and `cell.coordinate` once per cell in the worksheet iteration, pass them to all helper functions to bypass duplicate property/attribute access, and introduce a local `_html_escape_cache` to cache HTML-escaped text.
