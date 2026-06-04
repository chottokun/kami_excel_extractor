## 2025-05-15 - [Excel Style ID Caching]
**Learning:** `openpyxl` reuses `style_id` for cells with identical formatting. Caching style-related string generation (CSS, units, borders) per `style_id` significantly reduces CPU time in large sheets by avoiding redundant property access and string concatenations. In benchmarks, this yielded a ~40% performance gain for 50k cells.
**Action:** Always check for reusable ID-based objects in spreadsheet or document processing libraries before performing expensive property-to-string conversions in tight loops.
