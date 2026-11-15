# Excel

Use `xlsx-list-sheets` with an absolute `.xlsx` path to get exact worksheet names and visibility. Continue with `next_sheet` as `start_sheet`. Use `limit` for a smaller page of up to 100 worksheets. The list contains sheets with cells, including hidden worksheets.

Use `xlsx-read-range` with the exact `sheet` and an A1 address or rectangular `range` of at most 100 cells. The result includes blank cells at their original addresses. Formula cells distinguish formula text from the cached result. A null cached result indicates that no readable calculation result was stored in the workbook.

String values are limited to 160 characters per call. Continue a string by selecting that one cell and passing its `next_value_offset`, `next_formula_offset`, or `next_cached_value_offset` as `text_offset`. The offset applies to each string in that cell's result. Numeric and date values remain whole. Formulas use their stored results.

Example `{"path":"D:/work/budget.xlsx","sheet":"Summary","range":"B12","text_offset":160}`.
