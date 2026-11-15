[English](README.md)

# Vulcan Work

面向办公场景的 LuaSkills 技能包。首版提供 PDF 和 XLSX 的本地只读工具。

| 入口 | 返回内容 |
| --- | --- |
| `pdf-inspect` | 页数和交互式表单字段名 |
| `pdf-extract-text` | 指定范围内按页分组的文本 |
| `xlsx-list-sheets` | 工作表名称与可见状态 |
| `xlsx-read-range` | 单元格值、公式以及 Excel 已保存的公式结果 |

传入文件的绝对路径。PDF 页码从 1 开始。Excel 使用 A1 范围和准确的工作表名称。限制与字段说明见[工具帮助](help/help.md)。Word 和 PowerPoint 仍在规划中。

## 验证与打包

在开发用 Python 环境中安装 PyYAML（`python -m pip install pyyaml`）。

Windows 执行 `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify_skill.ps1`，Linux 或 macOS 执行 `bash scripts/verify_skill.sh`。调试器会准备固定版本的受管 Python 环境，并用合成文件调用实际工具。标准发布包输出到 `dist/`。

## 许可证与来源

项目采用 [MIT 许可证](LICENSE)。[第三方声明](licenses/THIRD_PARTY_NOTICES.md)记录了锁定的运行时依赖和许可证原文，以及设计阶段参考的 Anthropic Skills 来源。

项目基于 [LuaSkills demo-skill 模板](https://github.com/LuaSkills/demo-skill)初始化。
