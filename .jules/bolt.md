## 2025-05-15 - [Style Caching Optimization]
**Learning:** Caching style-dependent computations (CSS strings, unit info, borders) using 'cell.style_id' as the key significantly reduces redundant attribute lookups and method calls in 'openpyxl', leading to substantial performance gains in large workbooks.
**Action:** Always check for repeated property access on 'openpyxl' objects (like Font, Fill, Border) in tight loops and prefer caching results by 'style_id'.
