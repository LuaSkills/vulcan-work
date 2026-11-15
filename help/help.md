# Vulcan Work

Read local PDF, Excel, Word, and PowerPoint files. Pass absolute paths and request only the needed pages, rows, sections, or slides.

- PDF `pdf-inspect` discovers pages and form fields. `pdf-extract-text` reads specific pages.
- Excel `xlsx-list-sheets` discovers exact worksheet names. `xlsx-read-range` reads cells, formulas, and stored results.
- Word `docx-read-blocks` reads the main body. `docx-list-comments` reads comments by the IDs referenced in that body.
- PowerPoint `pptx-extract-text` reads slides in presentation order and can include speaker notes.

Item indices are one-based. A `next_*` item index resumes the list. Text offsets are zero-based Unicode character indices. To continue a truncated string, select that one item and pass the offset from its result. A null next index or offset marks the end.

The tools open the input files read-only. The maximum input size is 32 MiB. For Office ZIP files, the maximum expanded size is 128 MiB.
