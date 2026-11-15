# PDF

Use `pdf-inspect` with an absolute `.pdf` path to get `page_count`, `field_count`, and AcroForm field names. Use `start_field` and `limit` to page through names, up to 100 at a time. Continue with `next_field`. Form values are outside this tool's scope.

Use `pdf-extract-text` with `start_page` and `end_page`, inclusive and one-based. Select at most ten pages. Each page has up to 1,000 characters of text and a `next_text_offset`. For a continuation, set both page numbers to that page and pass `text_offset`.

Example `{"path":"D:/work/report.pdf","start_page":2,"end_page":2,"text_offset":1000}`.

Encrypted PDFs are unsupported. Text comes from the PDF text layer. Image-only pages may return an empty string.
