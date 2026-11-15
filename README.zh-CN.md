[English](README.md)

# Vulcan Work

Vulcan Work 是面向办公场景的 LuaSkills 技能包，计划支持处理 PDF、Excel、Word 和 PowerPoint 文件。

## 功能规划

- PDF · 查看文档信息、按页提取文本、列出可填写的表单字段。
- Excel（`.xlsx`）· 列出工作表、读取指定单元格范围，分别返回公式和缓存值。
- Word（`.docx`）· 读取文档结构、正文和批注。
- PowerPoint（`.pptx`）· 按幻灯片提取文本、生成预览。

首版从 PDF 和 Excel 的只读操作做起。当前仓库处于模板初始化阶段，上述办公工具仍在规划中。

## 许可证与来源

继承的模板采用 [MIT 许可证](LICENSE)，当前仓库所含组件的授权信息见[第三方声明](licenses/THIRD_PARTY_NOTICES.md)。后续如复用 Anthropic Skills 的代码或资源，发布包会随附适用的许可证和来源说明。

本仓库 fork 自 [LuaSkills/demo-skill](https://github.com/LuaSkills/demo-skill)，开发和打包流程见上游模板文档。
