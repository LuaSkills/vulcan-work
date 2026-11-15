# PowerPoint

Use `pptx-extract-text` with an absolute `.pptx` path. It reads up to ten slides in presentation order, including text boxes, grouped shapes, and table cells. It reports each slide's stored hidden status. Use `next_slide` as `start_slide`.

Set `include_notes` to true to read the speaker-notes body separately. Slide text and notes each have a 1,000-character budget. To continue either one, set `limit` to 1 and pass the corresponding `next_text_offset` as `text_offset` for slide text or `notes_offset` for notes.

Example `{"path":"D:/work/briefing.pptx","start_slide":4,"limit":1,"include_notes":true}`.

This reads stored text. Slide master content, images, and rendered previews are outside this tool's scope.

## Replace slide text

Use `pptx-replace-text` with a source `path`, an unused absolute `output_path`, and the one-based `slide` position. Provide exact case-sensitive `find` and `replacement` strings on a single line. Set `expected_matches` to the exact number of occurrences you intend to replace. A mismatch leaves the output path unused.

A match can span styled runs in the same paragraph. The replacement inherits the first matched run's styling. Other slides and notes stay in their existing package parts.

Example `{"path":"D:/work/review.pptx","output_path":"D:/work/review-approved.pptx","slide":2,"find":"Draft","replacement":"Approved","expected_matches":1}`.
