## 2025-06-01 - [Caching openpyxl style computations]
**Learning:** In openpyxl, many cells share the same `style_id`. Caching results of expensive style-related computations (like CSS string generation, border info parsing, and unit inference) based on this ID significantly improves performance for large workbooks.
**Action:** Always check if a cell's styling can be cached via `style_id` when performing repetitive visual extraction.
