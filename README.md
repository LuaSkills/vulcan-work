[简体中文](README.zh-CN.md)

# Vulcan Work

A LuaSkills package for reading local office files. This first version supports PDF and XLSX.

| Entry | Result |
| --- | --- |
| `pdf-inspect` | Page count and interactive form field names |
| `pdf-extract-text` | Text grouped by page for a requested range |
| `xlsx-list-sheets` | Worksheet names and visibility |
| `xlsx-read-range` | Cell values, formulas, and any results already saved by Excel |

Supply an absolute file path. PDF page numbers are one-based. Excel ranges use A1 notation and the exact worksheet name. See [tool help](help/help.md) for limits and result details. Word and PowerPoint are planned.

## Verify and package

Install PyYAML in the development Python environment (`python -m pip install pyyaml`).

On Windows, run `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify_skill.ps1`. On Linux or macOS, run `bash scripts/verify_skill.sh`. The debugger installs the pinned managed Python environment and runs the tools against synthetic files. The standard package is written to `dist/`.

## License and acknowledgments

The project uses the [MIT License](LICENSE). The [third-party notices](licenses/THIRD_PARTY_NOTICES.md) record the locked runtime dependencies and their license texts, as well as the Anthropic Skills material referenced during design.

The project started from the [LuaSkills demo-skill template](https://github.com/LuaSkills/demo-skill).
