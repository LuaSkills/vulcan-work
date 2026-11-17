# PowerPoint

Read this topic to choose between stored slide text, speaker notes, editable slides, and a rendered preview. When editing or removing existing slides, read the relevant slide first if its position or text needs confirmation.

Use `vulcan_work_pptx_extract_text` with an absolute `.pptx` path. It reads up to ten slides in presentation order, including text boxes, grouped shapes, and table cells. It reports each slide's stored hidden status. Use `next_slide` as `start_slide`.

Set `include_notes` to true to read the speaker-notes body separately. Slide text and notes each have a 1,000-character budget. To continue either one, set `limit` to 1 and pass the corresponding `next_text_offset` as `text_offset` for slide text or `notes_offset` for notes.

Example `{"path":"D:/work/briefing.pptx","start_slide":4,"limit":1,"include_notes":true}`.

The extraction entry reads stored text. For appearance, layout, or a PDF preview, use `vulcan_work_pptx_export_preview` and inspect its PDF output.

## Replace slide text

Use `vulcan_work_pptx_replace_text` with a source `path`, an unused absolute `output_path`, and the one-based `slide` position. Provide exact case-sensitive `find` and `replacement` strings on a single line. Set `expected_matches` to the exact number of occurrences you intend to replace. A mismatch leaves the output path unused.

A match can span styled runs in the same paragraph. The replacement inherits the first matched run's styling. Other slides and notes stay in their existing package parts.

Example `{"path":"D:/work/review.pptx","output_path":"D:/work/review-approved.pptx","slide":2,"find":"Draft","replacement":"Approved","expected_matches":1}`.

## Manage slides and render previews

`vulcan_work_pptx_create` generates an editable presentation from up to 30 `slides` with a title and optional body paragraphs. `vulcan_work_pptx_append_slides` adds such slides to an existing deck using its first slide's layout or the deck's first master layout. These are plain title-and-body slides.

`vulcan_work_pptx_remove_slides` deletes one-based positions from a copy. Slide parts and uniquely owned notes or media are removed when they are no longer reachable. Keep at least one slide in the copy. The tool also rejects a removal when another part still links to the slide. `vulcan_work_pptx_export_preview` renders the presentation to a new PDF in presentation order.

Creation and rendered previews require LibreOffice Impress. Use the preview PDF to review the slide visuals. LibreOffice may export Chinese text in a different reading order from the source, so use the source PPTX for text extraction and indexing. Editing, appending, and removing slides work directly on the package. Every write requires an unused output path. To verify a text edit or slide count, extract the relevant text or slide metadata from the output copy.

## Check LibreOffice before creating or rendering

Before calling `vulcan_work_pptx_create` or `vulcan_work_pptx_export_preview`, check for LibreOffice Impress on the host running Vulcan Code. Look for `soffice` on that host's PATH. For a desktop installation, also check `%ProgramFiles%/LibreOffice/program/soffice.com` on Windows or `/Applications/LibreOffice.app/Contents/MacOS/soffice` on macOS.

If LibreOffice is unavailable, tell the user which requested operation needs it and ask whether they want to install it. Wait for their approval before installing. If they choose to continue without LibreOffice, let them know that creation and PDF preview are unavailable. For an existing presentation, ask the user to open it in their presentation app to check the layout and appearance.

## Entry arguments

Pass the fields shown in each example as separate entry arguments. Treat the example's `input` heading as the title of the whole argument list. For a call, pass one named argument for each direct child of that heading. Keep nested lists or objects inside their own named field.
