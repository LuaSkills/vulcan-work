# Third-party notices

Development tooling originated in the [LuaSkills demo-skill template](https://github.com/LuaSkills/demo-skill) under the project MIT license.

The managed Python environment installs the exact dependencies recorded in
[`python/requirements.lock`](../python/requirements.lock). The archive includes
the following license texts.

| Dependency | Version | License | Source |
| --- | --- | --- | --- |
| pypdf | 7.1.0 | [BSD-3-Clause](pypdf-LICENSE.txt) | [Release](https://pypi.org/project/pypdf/7.1.0/) |
| openpyxl | 3.1.5 | [MIT](openpyxl-LICENCE.rst) | [Release](https://pypi.org/project/openpyxl/3.1.5/) |
| et-xmlfile | 2.0.0 | [MIT and Python license notices](et-xmlfile-LICENCE.rst) ([Python license text](et-xmlfile-LICENCE.python)) | [Release](https://pypi.org/project/et-xmlfile/2.0.0/) |
| defusedxml | 0.7.1 | [Python Software Foundation License](defusedxml-LICENSE.txt) | [Release](https://pypi.org/project/defusedxml/0.7.1/) |

The PDF, spreadsheet, Word, and PowerPoint workflows were informed by
[Anthropic Skills at commit 1780c5e](https://github.com/anthropics/skills/tree/1780c5e5a552fd4d1c9d6a7b0cab26b0c2bef064/skills).
The current package uses original handlers and contains no Anthropic skill files
or assets.
