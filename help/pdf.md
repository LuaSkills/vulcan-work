# PDF

Use `pdf-inspect` with an absolute `.pdf` path to get `page_count`, `field_count`, and AcroForm field names. Use `start_field` and `limit` to page through names, up to 100 at a time. Continue with `next_field`. Form values are outside this tool's scope.

Use `pdf-extract-text` with `start_page` and `end_page`, inclusive and one-based. Select at most ten pages. Each page has up to 1,000 characters of text and a `next_text_offset`. For a continuation, set both page numbers to that page and pass `text_offset`.

Example `{"path":"D:/work/report.pdf","start_page":2,"end_page":2,"text_offset":1000}`.

Encrypted PDFs are unsupported. Text comes from the PDF text layer. Image-only pages may return an empty string.

## Write a PDF copy

Use `pdf-compose` with an ordered `sources` list and an unused absolute `output_path`. A source may specify an inclusive `start_page` and `end_page`; omit either endpoint to use the corresponding document boundary. Use one source for extraction and multiple sources for merging. Each call accepts up to ten inputs and 500 output pages.

Example `{"sources":[{"path":"D:/work/report.pdf","start_page":2,"end_page":4},{"path":"D:/work/appendix.pdf"}],"output_path":"D:/work/combined.pdf"}`.

Copying pages creates a new PDF. Existing digital signatures do not apply to the generated file.
