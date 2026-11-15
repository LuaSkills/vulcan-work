# Excel

Read this topic when the user asks for worksheet names, cell values or formulas, or wants to create or change a workbook. If an existing workbook's sheet name is unknown, list its worksheets first. Use the exact name returned for the range or edit.

Use `vulcan-work-xlsx-list-sheets` with an absolute `.xlsx` path to get exact worksheet names and visibility. Continue with `next_sheet` as `start_sheet`. Use `limit` for a smaller page of up to 100 worksheets. The list contains sheets with cells, including hidden worksheets.

Use `vulcan-work-xlsx-read-range` with the exact `sheet` and an A1 address or rectangular `range` of at most 100 cells. The result includes blank cells at their original addresses. Formula cells distinguish formula text from the cached result. A null cached result indicates that no readable calculation result was stored in the workbook.

String values are limited to 160 characters per call. Continue a string by selecting that one cell and passing its `next_value_offset`, `next_formula_offset`, or `next_cached_value_offset` as `text_offset`. The offset applies to each string in that cell's result. Numeric and date values remain whole. Formulas use their stored results.

Example `{"path":"D:/work/budget.xlsx","sheet":"Summary","range":"B12","text_offset":160}`.

To list sheets, pass `path` as its own argument. Example `{"path":"D:/work/budget.xlsx"}`.

## Write cells

Use `vulcan-work-xlsx-write-cells` with an unused absolute `output_path`, an exact `sheet` name, and up to 100 unique A1 cell updates. Set `path` to copy a supported workbook or omit it to create a new workbook with the named sheet. Each update accepts a scalar `value` or an explicit `formula`. A null value clears a cell. A string beginning with an equals sign stays literal when passed as `value`.

Example `{"output_path":"D:/work/summary.xlsx","sheet":"Summary","cells":[{"address":"A1","value":"Total"},{"address":"B1","formula":"=SUM(B2:B10)"}]}`.

Saving a workbook copy clears stored formula results. A spreadsheet application can recalculate them when the copy is opened. Editing copies unused parts as-is, including charts, drawings, layouts, styles, and unrelated resources. Signed files and targeted grouped formulas, rich-value metadata, and merged non-anchor cells are rejected. When a value changes, formula caches are cleared across the workbook and Excel is asked to recalculate.

After a write, use `vulcan-work-xlsx-read-range` on the output copy to verify the affected addresses. For formulas, verify the formula text; a newly written formula may have no saved result until Excel recalculates it.

## Entry arguments

Pass the fields shown in each example as separate entry arguments. Treat the schema's `input` heading as the title of the whole argument list. For a call, pass one named argument for each direct child of that heading. Keep nested lists or objects inside their own named field.
