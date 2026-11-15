[简体中文](README.zh-CN.md)

# Vulcan Work

A LuaSkills package for reading local office files. Supported formats are PDF, XLSX, DOCX, and PPTX.

| Entry | Result |
| --- | --- |
| `pdf-inspect` | Page count and interactive form field names |
| `pdf-extract-text` | Text grouped by page for a requested range |
| `xlsx-list-sheets` | Worksheet names and visibility |
| `xlsx-read-range` | Cell values, formulas, and any results already saved by Excel |
| `docx-read-blocks` | Main-body paragraphs and tables, paragraph styles, and comment references |
| `docx-list-comments` | Comment IDs, authors, dates, and text |
| `pptx-extract-text` | Text by slide in presentation order, including groups and tables |

Supply an absolute file path. PDF page numbers are one-based. Excel ranges use A1 notation and the exact worksheet name. See [tool help](help/help.md) for limits and result details. Long text can be continued by reading the same item with its returned offset. PowerPoint can include speaker notes.

## Verify and package

Install PyYAML in the development Python environment (`python -m pip install pyyaml`).

On Windows, run `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify_skill.ps1`. On Linux or macOS, run `bash scripts/verify_skill.sh`. The debugger installs the pinned managed Python environment and runs the tools against synthetic files. The standard package is written to `dist/`.

## License and acknowledgments

The project uses the [MIT License](LICENSE). The [third-party notices](licenses/THIRD_PARTY_NOTICES.md) record the locked runtime dependencies and their license texts, as well as the Anthropic Skills material referenced during design.
