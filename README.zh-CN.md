[English](README.md)

# Vulcan Work

面向办公场景的 LuaSkills 技能包。提供 PDF、XLSX、DOCX 和 PPTX 的本地读写工具。

| 入口 | 返回内容 |
| --- | --- |
| `pdf-inspect` | 页数和交互式表单字段名 |
| `pdf-extract-text` | 指定范围内按页分组的文本 |
| `xlsx-list-sheets` | 工作表名称与可见状态 |
| `xlsx-read-range` | 单元格值、公式以及 Excel 已保存的公式结果 |
| `docx-read-blocks` | 正文段落、表格、段落样式和批注引用 |
| `docx-list-comments` | 批注标识、作者、日期和文本 |
| `pptx-extract-text` | 按放映顺序提取文本框、组合形状和表格中的文本 |
| `pdf-compose` | 合并 PDF，或将指定页提取到新 PDF |
| `xlsx-write-cells` | 新建工作簿，或向工作簿副本批量写入单元格 |
| `docx-append-paragraphs` | 新建 Word 文档，或向副本追加段落和标题 |
| `pptx-replace-text` | 在新副本的指定幻灯片上精确替换文字 |

传入文件的绝对路径。PDF 页码从 1 开始。Excel 使用 A1 范围和准确的工作表名称。限制与字段说明见[工具帮助](help/help.md)。长文本可按返回的偏移续读。PowerPoint 可同时读取演讲者备注。

写入时须提供已有目录中的新文件绝对路径。操作保留源文件和已有输出文件。

## 验证与打包

在开发用 Python 环境中安装 PyYAML（`python -m pip install pyyaml`）。

Windows 执行 `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify_skill.ps1`，Linux 或 macOS 执行 `bash scripts/verify_skill.sh`。调试器会准备固定版本的受管 Python 环境，并用合成文件调用实际工具。标准发布包输出到 `dist/`。

## 许可证与来源

项目采用 [MIT 许可证](LICENSE)。[第三方声明](licenses/THIRD_PARTY_NOTICES.md)记录了锁定的运行时依赖和许可证原文，以及设计阶段参考的 Anthropic Skills 来源。
