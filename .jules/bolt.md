
## 2026-06-18 - [Excel Style Caching]
**Learning:** Accessing `openpyxl` cell attributes (fill, border, font, number_format) is extremely expensive in large loops. Since Excel cells often share the same style, these attributes can be cached using `cell.style_id` as a key.
**Action:** Use a per-worksheet style cache keyed by `style_id` to store pre-calculated CSS strings and metadata, significantly reducing redundant attribute lookups.
