---
name: office-workflow
description: For local PDF, Excel, Word, and PowerPoint tasks, load this guide before using generic tools. Covers PDF page counts and forms, spreadsheets, documents, and slides.
---

Use `tool_call` with this skill name, the public action name below, and a JSON `arguments` object. For example, to inspect the attached PDF at its resolved local path:

```json
{"skill":"luaskills-vulcan-work:office-workflow","action":"vulcan_work_pdf_inspect","arguments":{"path":"D:/work/report.pdf"}}
```

Use the actual absolute path to the input file. Put changes in a separate `output_path` unless the user has selected a permitted destination for the copy. Page, block, slide, row and column indices start at 1.

| Tool | Required arguments and use |
| --- | --- |
| `vulcan_work_pdf_inspect` | `path`. For stored form values, add `include_values=true`. |
| `vulcan_work_pdf_extract_text` | `path`, `start_page`, `end_page`. Use `text_offset` to continue an output page. |
| `vulcan_work_pdf_ocr_pages` | `path`, `start_page`, `end_page`. Use for scans. Optional `language` specifies OCR language. |
| `vulcan_work_pdf_compose` | `sources`, `output_path`. Each source specifies a PDF path and optionally a page range. |
| `vulcan_work_pdf_fill_form` | `path`, `output_path`, `fields`. Inspect field names in the source before filling them. |
| `vulcan_work_xlsx_list_sheets` | `path`. Optional `start_sheet` and `limit` advance the listing. |
| `vulcan_work_xlsx_read_range` | `path`, `sheet`, `range` such as `A1:C12`. Read existing formula results when available. |
| `vulcan_work_xlsx_write_cells` | `sheet`, `cells`, `output_path`. Pass `path` to edit a copy or omit it to create a workbook. |
| `vulcan_work_docx_read_blocks` | `path`. Use `start_block` or `limit` to read the next blocks. |
| `vulcan_work_docx_list_comments` | `path`. Optional `start_comment` and `limit` advance the listing. |
| `vulcan_work_docx_append_paragraphs` | `paragraphs`, `output_path`. Pass `path` to append to a copy of an existing document. |
| `vulcan_work_docx_replace_text` | `path`, `output_path`, `block`, `find`, `replacement`, `expected_matches`. For table cells, also pass `row` and `column`. |
| `vulcan_work_pptx_extract_text` | `path`. Add `include_notes=true` to read speaker notes. |
| `vulcan_work_pptx_replace_text` | `path`, `output_path`, `slide`, `find`, `replacement`, `expected_matches`. |
| `vulcan_work_pptx_create` | `output_path`, `slides`. |
| `vulcan_work_pptx_append_slides` | `path`, `output_path`, `slides`. |
| `vulcan_work_pptx_remove_slides` | `path`, `output_path`, `slides` as an array of slide numbers. |
| `vulcan_work_pptx_export_preview` | `path`, `output_path` for the generated PDF preview. |

Read the relevant reference for nested arguments, output limits, or file-specific workflows: [PDF](references/pdf.md), [Excel](references/xlsx.md), [Word](references/docx.md), [PowerPoint](references/pptx.md).
