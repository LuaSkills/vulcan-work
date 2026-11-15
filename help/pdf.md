# PDF

Use `pdf-inspect` with an absolute `.pdf` path to get `page_count`, `field_count`, and AcroForm field names. Use `start_field` and `limit` to page through names, up to 100 at a time. Continue with `next_field`. Form values are outside this tool's scope.

Use `pdf-extract-text` with `start_page` and `end_page`, inclusive and one-based. Select at most ten pages. Each page has up to 1,000 characters of text and a `next_text_offset`. For a continuation, set both page numbers to that page and pass `text_offset`.

Example `{"path":"D:/work/report.pdf","start_page":2,"end_page":2,"text_offset":1000}`.

Encrypted PDFs are unsupported. Text comes from the PDF text layer. Image-only pages may return an empty string.

## Write a PDF copy

Use `pdf-compose` with an ordered `sources` list and an unused absolute `output_path`. A source may specify an inclusive `start_page` and `end_page`; omit either endpoint to use the corresponding document boundary. Use one source for extraction and multiple sources for merging. Each call accepts up to ten inputs and 500 output pages.

Example `{"sources":[{"path":"D:/work/report.pdf","start_page":2,"end_page":4},{"path":"D:/work/appendix.pdf"}],"output_path":"D:/work/combined.pdf"}`.

Copying pages creates a new PDF. Existing digital signatures do not apply to the generated file.

## Fill or recognize a form

Use `pdf-fill-form` with a source PDF, a new `output_path`, and up to 100 pairs of `name` and `value`. `pdf-inspect` returns field types and available button or choice states. Set `include_values` only when the current values are needed. Text, choice, checkbox and radio fields are supported in standard AcroForms. XFA, signed PDFs, signature fields, and read-only fields are rejected. The writer verifies the stored value in the generated PDF.

Use `pdf-ocr-pages` when the original page is scanned or image-based. Provide an inclusive range of up to three pages. OCR text uses the same `next_text_offset` as ordinary PDF extraction. The host needs Poppler's `pdftoppm` on PATH and an installed OCR engine. With Tesseract, `auto` selects English; install `chi_sim` or `chi_tra` to request Chinese. On Windows, when Tesseract is absent, recognition uses the Windows OCR engine and its installed languages.
