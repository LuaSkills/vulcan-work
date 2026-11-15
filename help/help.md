# Vulcan Work

Use Vulcan Work when the user asks to inspect, extract from, edit, create, or render an actual local PDF, XLSX, DOCX, or PPTX file. Use the file's format and the requested result to choose an entry. A general explanation or a draft for the chat does not require a file tool. For legacy Office formats, ask for a supported file or an explicit conversion request.

## Choose the help and entry

Read the matching format topic when deciding between operations or planning a multi-step task. Read the selected entry's own help and argument schema before its first call, even after consulting the format topic. Repeat this for each new entry in a multi-step task. The format topic explains which metadata to discover first, how to resume a partial result, and how to verify a new file.

| Topic | User intent | Entry |
| --- | --- | --- |
| `pdf` | Count pages, discover form fields, or read stored field values | `vulcan-work-pdf-inspect` |
| `pdf` | Read text on specified pages | `vulcan-work-pdf-extract-text` |
| `pdf` | Recognize text on scanned or image-only pages | `vulcan-work-pdf-ocr-pages` |
| `pdf` | Merge files or copy selected pages into a new PDF | `vulcan-work-pdf-compose` |
| `pdf` | Fill an interactive form in a new PDF copy | `vulcan-work-pdf-fill-form` |
| `xlsx` | Find exact worksheet names, including hidden sheets | `vulcan-work-xlsx-list-sheets` |
| `xlsx` | Read cells, formulas, or saved formula results | `vulcan-work-xlsx-read-range` |
| `xlsx` | Create a workbook or change cells in a copy | `vulcan-work-xlsx-write-cells` |
| `docx` | Read main-body paragraphs and tables | `vulcan-work-docx-read-blocks` |
| `docx` | Read embedded review comments | `vulcan-work-docx-list-comments` |
| `docx` | Create a document or append headings and paragraphs | `vulcan-work-docx-append-paragraphs` |
| `docx` | Replace precise text in one existing paragraph or table cell | `vulcan-work-docx-replace-text` |
| `pptx` | Read stored slide text or speaker notes | `vulcan-work-pptx-extract-text` |
| `pptx` | Create an editable presentation | `vulcan-work-pptx-create` |
| `pptx` | Add slides to a copy | `vulcan-work-pptx-append-slides` |
| `pptx` | Remove selected slides from a copy | `vulcan-work-pptx-remove-slides` |
| `pptx` | Replace precise text on a selected slide in a copy | `vulcan-work-pptx-replace-text` |
| `pptx` | Render slide visuals as a new PDF preview | `vulcan-work-pptx-export-preview` |

## Work with the result

Pass the fields shown in each example as separate entry arguments. Treat the schema's `input` heading as the title of the whole argument list. For a call, pass one named argument for each direct child of that heading. Keep nested lists or objects inside their own named field.

Use absolute paths. Use the smallest page, cell, block, or slide range that answers the request. Item positions are one-based. A non-null `next_*` item position resumes a list. A text offset counts Unicode characters from zero and resumes one selected item. Follow the topic for the exact argument name.

For an edit, inspect the source first when a field, sheet, block, or slide must be identified. Write to an unused absolute `output_path` in an existing directory and then inspect the output when verification is needed. Run dependent steps in order and check that each step succeeded before using its result. If a write is denied, leave the output untouched. Each source is opened read-only. Input files are limited to 32 MiB, and Office ZIP files to 128 MiB expanded.
