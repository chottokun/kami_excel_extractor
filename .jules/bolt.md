## 2025-05-15 - [MetadataExtractor] Optimization of Cell Style Attribute Access

**Learning:** Accessing cell font properties (like `cell.font.b`) in `openpyxl` during large sheet iteration is expensive due to repeated proxy object creation and lookups.
**Action:** Use a dedicated cache (`self._bold_cache`) keyed by `cell.style_id` to store font weight status. Combine this with pre-calculating other style-related strings (`style_str`, `unit`) once per cell to avoid redundant method calls in the extraction loop.
