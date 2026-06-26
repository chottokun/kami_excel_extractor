## 2025-05-15 - [Style Caching in MetadataExtractor]
**Learning:** In Excel files with many cells sharing the same style, accessing style properties (`cell.font`, `cell.border`, `cell.number_format`) repeatedly via `openpyxl` is expensive. `openpyxl` assigns a `style_id` to each unique combination of style attributes.
**Action:** Implement a per-file style cache keyed by `cell.style_id` to store pre-calculated CSS strings, border info, and unit info. This significantly reduces the cumulative time spent in style extraction methods during large sheet processing.
