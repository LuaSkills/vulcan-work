# PDF

Read this topic when the task involves PDF pages, text, scans, or form fields. Check page and field metadata first when it is needed to choose a page range or address a field.

## Inspect pages and forms

Use `vulcan_work_pdf_inspect` with an absolute `.pdf` path to get `page_count`, `field_count`, and AcroForm field names, types, read-only flags, and available states. The default response omits stored values. Pass `include_values: true` when the user asks for current values or when verifying a filled copy. The `fields[].value` member can be a string, a list, or null. Use `start_field` and `limit` to page through up to 100 fields at a time, and resume with `next_field`.

Example `{"path":"D:/work/form.pdf","include_values":true}`.

## Read page text

Use `vulcan_work_pdf_extract_text` with inclusive, one-based `start_page` and `end_page`. Select at most ten pages. Each page has up to 1,000 characters of text. Continue an individual page by setting both page numbers to that page and passing its `next_text_offset` as `text_offset`.

Example `{"path":"D:/work/report.pdf","start_page":2,"end_page":2,"text_offset":1000}`.

Use `vulcan_work_pdf_ocr_pages` when the user identifies a scan or when the needed page is image-based and text extraction does not return its text. OCR accepts an inclusive range of up to three pages and uses the same text-offset convention. Install Poppler and make `pdftoppm` available on the Vulcan Code host's PATH on every platform, including Windows. It rasterizes PDF pages before a local OCR engine recognizes them. With Tesseract, `auto` selects English. Install `chi_sim` or `chi_tra` to request Chinese. On Windows, when Tesseract is absent, recognition uses Windows OCR and its installed languages. Encrypted PDFs are unsupported.

## Write a PDF copy

Use `vulcan_work_pdf_compose` with an ordered `sources` list and an unused absolute `output_path`. Each source may specify inclusive `start_page` and `end_page` values. An omitted endpoint uses the corresponding document boundary. Use one source to extract pages and multiple sources to merge them. Each call accepts up to ten inputs and 500 output pages. Inspect the result if the page count needs verification. Existing digital signatures do not apply to the generated file.

Example `{"sources":[{"path":"D:/work/report.pdf","start_page":2,"end_page":4},{"path":"D:/work/appendix.pdf"}],"output_path":"D:/work/combined.pdf"}`.

## Fill an interactive form

First use `vulcan_work_pdf_inspect` to identify each target by its exact field name and check its type, read-only flag, and accepted states. Then use `vulcan_work_pdf_fill_form` with up to 100 pairs of `name` and `value`, the source path, and a new `output_path`. Checkbox, radio, and choice updates must use the exact state from the field metadata. Text, choice, checkbox, and radio fields are supported in standard AcroForms. XFA, signed PDFs, signature fields, and read-only fields are rejected. The writer verifies the stored values. To report the final values to the user, inspect the output copy with `include_values: true`.

Example `{"path":"D:/work/form.pdf","output_path":"D:/work/form-filled.pdf","fields":[{"name":"Department","value":"Training"}]}`. For this call the three separate argument names are `path`, `output_path`, and `fields`.

## Entry arguments

Pass the fields shown in each example as separate entry arguments. Treat the example's `input` heading as the title of the whole argument list. For a call, pass one named argument for each direct child of that heading. Keep nested lists or objects inside their own named field.
