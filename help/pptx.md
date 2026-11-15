# PowerPoint

Use `pptx-extract-text` with an absolute `.pptx` path. It reads up to ten slides in presentation order, including text boxes, grouped shapes, and table cells. It reports each slide's stored hidden status. Use `next_slide` as `start_slide`.

Set `include_notes` to true to read the speaker-notes body separately. Slide text and notes each have a 1,000-character budget. To continue either one, set `limit` to 1 and pass the corresponding `next_text_offset` as `text_offset` for slide text or `notes_offset` for notes.

Example `{"path":"D:/work/briefing.pptx","start_slide":4,"limit":1,"include_notes":true}`.

This reads stored text. Slide master content, images, and rendered previews are outside this tool's scope.
