# Word

Use `docx-read-blocks` with an absolute `.docx` path. It lists main-body paragraphs and tables in order, with style IDs and direct outline levels on paragraphs. Inside a table, tabs separate cells and newlines separate rows. `comment_ids` contains comment identifiers referenced in a block.

Use `docx-list-comments` to read embedded comment bodies and their author and date metadata. Comment IDs match `comment_ids` from the body. The comment index is its one-based position in this list.

Each tool returns up to 20 items at a time. Use `next_block` as `start_block` or `next_comment` as `start_comment`. To read past the 1,000-character limit of one block or comment, set `limit` to 1 and pass its `next_text_offset` as `text_offset`.

Example `{"path":"D:/work/notes.docx","start_block":3,"limit":1,"text_offset":1000}`.
