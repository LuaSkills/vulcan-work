# Vulcan Work

Use these read-only tools for local office files. Supply absolute file paths and narrow selections.

- `pdf-inspect` returns the page count and up to 100 AcroForm field names. Password-protected and encrypted PDFs are unsupported.
- `pdf-extract-text` accepts an inclusive one-based range of at most ten pages. It returns each requested page separately, with up to 1,000 text characters per page and an explicit truncation flag. An empty text result can occur with image-only pages. The tool does not perform OCR.
- `xlsx-list-sheets` lists worksheets with their visibility and reports the total when the list is capped.
- `xlsx-read-range` accepts a single A1 cell or a rectangular range of up to 100 cells. Use the worksheet's exact name. A formula cell includes the formula and its saved result. A null saved result means the workbook has no readable cached result for that cell. Text values longer than 160 characters are truncated with a flag.

The input file remains read-only. Files larger than 32 MiB are rejected. Excel recalculation, document editing, and remote URLs are outside the current tool set.
