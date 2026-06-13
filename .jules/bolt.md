## 2025-05-15 - [Optimize Excel style extraction with caching]
**Learning:** Redundant style lookups in openpyxl (borders, fonts, fills) for each cell are a major bottleneck in large spreadsheets. openpyxl provides a `style_id` that can be used to cache these properties per workbook.
**Action:** Always implement a per-workbook style cache keyed by `style_id` when processing large Excel files to avoid O(N*M) redundant CSS/metadata generation.
