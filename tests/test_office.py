"""Exercise the actual LuaSkills entries against small synthetic office files.
使用小型合成办公文件测试实际的 LuaSkills 入口。
"""

import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree
from zipfile import ZipFile

# Anchor paths to the repository regardless of the test runner's working directory.
# 无论测试从哪个工作目录启动，均使用仓库作为路径基准。
ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"


def call_tool(entry: str, arguments: dict) -> subprocess.CompletedProcess:
    """Invoke entry with arguments through the repository's platform debugger.
    通过仓库的平台调试器，使用 arguments 调用 entry。
    """
    with TemporaryDirectory() as directory:
        # An args file preserves exact path spelling and Unicode across shells.
        # 使用参数文件在不同 shell 中保留路径拼写和 Unicode。
        args_file = Path(directory) / "arguments.json"
        args_file.write_text(json.dumps(arguments), encoding="utf-8")
        if sys.platform == "win32":
            command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                       str(ROOT / "scripts/debug_skill.ps1"), "-Command", "call", "-Tool",
                       "vulcan-work-" + entry, "-ArgsFile", str(args_file), "-Output", "json"]
        else:
            command = ["bash", str(ROOT / "scripts/debug_skill.sh"), "call", "--tool",
                       "vulcan-work-" + entry, "--args-file", str(args_file), "--output", "json"]
        return subprocess.run(command, cwd=ROOT, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=180, check=False)


def replace_fixture_xml(source: Path, target: Path, member_name: str, old: str, new: str) -> None:
    """Copy a synthetic Office ZIP, replacing one known XML text value.
    复制合成 Office ZIP 并替换一个已知的 XML 文本值。
    """
    # Encode replacement text so the fixture can exercise Unicode character offsets.
    # 编码替换文本，使样例能够验证 Unicode 字符偏移。
    before, after = old.encode("utf-8"), new.encode("utf-8")
    with ZipFile(source) as original, ZipFile(target, "w") as modified:
        for name in original.namelist():
            # Preserve every unrelated package member.
            # 保留归档中其他成员。
            content = original.read(name)
            if name == member_name:
                if before not in content:
                    raise ValueError(f"Fixture XML changed: {member_name}.")
                content = content.replace(before, after, 1)
            modified.writestr(name, content)


def write_pdf_fixture(path: Path, text: str) -> None:
    """Write a minimal one-page PDF with an ASCII text layer and valid cross-reference offsets.
    写入带有 ASCII 文本层和有效交叉引用偏移的最小单页 PDF。
    """
    # The synthetic page draws one literal string using a PDF standard font.
    # 合成页面使用 PDF 标准字体绘制一个字面字符串。
    stream = f"BT /F1 12 Tf 10 220 Td ({text}) Tj ET".encode("ascii")
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 240] "
         b"/Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream",
    ]
    # Record each byte offset while assembling the PDF objects.
    # 写入 PDF 对象时记录每个对象的字节偏移。
    body = bytearray(b"%PDF-1.4\n")
    offsets = []
    for number, obj in enumerate(objects, 1):
        offsets.append(len(body))
        body.extend(f"{number} 0 obj\n".encode("ascii") + obj + b"\nendobj\n")
    # The trailer points at the exact cross-reference table appended below.
    # 文件尾部指向下方追加的精确交叉引用表。
    xref = len(body)
    body.extend(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode("ascii"))
    for offset in offsets:
        body.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    body.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode("ascii"))
    path.write_bytes(body)


class OfficeEntryTests(unittest.TestCase):
    """Validate manifest-driven calls and their user-visible output contracts.
    验证清单驱动的调用及其面向用户的输出约定。
    """

    def result(self, entry: str, filename: str | Path, **arguments: object) -> dict:
        """Return the decoded user-visible result for a fixture and operation.
        返回指定样例与操作的用户可见结果。
        """
        return self.invoke(entry, path=str(FIXTURES / filename), **arguments)

    def invoke(self, entry: str, **arguments: object) -> dict:
        """Return the decoded result for a successful call with explicit arguments.
        返回显式参数调用成功后的解码结果。
        """
        # The debugger may print bootstrap status before its JSON envelope.
        # 调试器可能在 JSON 外层结构之前打印启动状态。
        finished = call_tool(entry, arguments)
        self.assertEqual(finished.returncode, 0, finished.stdout + finished.stderr)
        # JSON envelope starts after any setup lines and contains a serialized tool response.
        # JSON 外层结构位于初始化日志之后，并包含序列化的工具结果。
        envelope = json.loads(finished.stdout[finished.stdout.index("{\n"):])
        return json.loads(envelope["result"]["content"])

    def test_pdf_inspect(self) -> None:
        """Confirm the tool returns page counts and interactive form names.
        确认工具返回页数和交互表单字段名。
        """
        # The synthetic PDF has one page and one form field.
        # 合成 PDF 只有一页和一个表单字段。
        result = self.result("pdf-inspect", "office.pdf")
        self.assertEqual(result["page_count"], 1)
        self.assertEqual(result["field_names"], ["Department"])

    def test_pdf_extract_text(self) -> None:
        """Confirm one-based selection extracts actual page text.
        确认一基范围能够提取实际页面文本。
        """
        # Request the first page through the same schema exposed to the host.
        # 使用暴露给宿主的同一套 Schema 请求第一页。
        result = self.result("pdf-extract-text", "office.pdf", start_page=1, end_page=1)
        self.assertEqual(result["pages"][0]["page"], 1)
        self.assertIn("Vulcan Work", result["pages"][0]["text"])

    def test_xlsx_list_sheets(self) -> None:
        """Confirm sheet identity and hidden state are preserved.
        确认工作表名称与隐藏状态得到保留。
        """
        # Hidden is a worksheet state and must be represented explicitly.
        # 隐藏属于工作表状态，需要显式返回。
        result = self.result("xlsx-list-sheets", "office.xlsx")
        self.assertIn({"name": "Hidden", "visibility": "hidden"}, result["sheets"])
        self.assertEqual(result["sheet_count"], 2)

    def test_xlsx_range_contains_values_formulas_and_caches(self) -> None:
        """Confirm cached and uncached formulas, values, and blanks are distinguishable.
        确认可区分有缓存和无缓存的公式、普通值以及空白单元格。
        """
        # Include a blank column and row outside the recorded worksheet dimensions.
        # 包含一列和一行超出工作表记录维度的空白单元格。
        result = self.result("xlsx-read-range", "office.xlsx", sheet="Data", range="A1:E2")
        # Resolve returned cells by address rather than relying on their position.
        # 按地址定位返回单元格，无需依赖其位置。
        cells = {cell["address"]: cell for cell in result["cells"]}
        self.assertEqual(cells["A1"]["value"], "Label")
        self.assertEqual(cells["B1"]["value"], 6)
        self.assertEqual(cells["C1"]["formula"], "=B1*7")
        self.assertEqual(cells["C1"]["cached_value"], 42)
        self.assertEqual(cells["D1"]["kind"], "formula")
        self.assertIsNone(cells["D1"]["cached_value"])
        self.assertIsNone(cells["E1"]["value"])
        self.assertIsNone(cells["A2"]["value"])
        self.assertIsNone(cells["E2"]["value"])
        self.assertEqual(len(cells), 10)

    def test_out_of_bounds_pdf_range_fails(self) -> None:
        """Confirm invalid page selection surfaces an actionable error.
        确认无效页码范围会返回可操作的错误。
        """
        # There is no page two in the fixture.
        # 样例中没有第二页。
        finished = call_tool("pdf-extract-text", {
            "path": str(FIXTURES / "office.pdf"), "start_page": 2, "end_page": 2,
        })
        self.assertNotEqual(finished.returncode, 0)
        self.assertIn("document has 1 pages", finished.stdout + finished.stderr)


    def test_word_blocks_keep_structure_and_comments(self) -> None:
        """Read a heading, a paragraph with a comment, and a table in document order.
        按文档顺序读取标题、带批注的段落和表格。
        """
        # The first page ends immediately after the heading.
        # 第一页仅包含标题。
        first = self.result("docx-read-blocks", "office.docx", limit=1)
        self.assertEqual(first["block_count"], 3)
        self.assertEqual(first["next_block"], 2)
        self.assertEqual(first["blocks"][0]["style_id"], "Heading1")
        # The rest demonstrates tab separators, comment anchors, and table cell order.
        # 剩余内容验证制表符、批注锚点和表格单元格顺序。
        rest = self.result("docx-read-blocks", "office.docx", start_block=first["next_block"])
        paragraph, table = rest["blocks"]
        self.assertEqual(paragraph["text"], "Review\tneeded")
        self.assertEqual(len(paragraph["comment_ids"]), 1)
        self.assertEqual(table["kind"], "table")
        self.assertEqual(table["text"], "Item\tTotal")
        self.assertIsNone(rest["next_block"])

    def test_word_comments_return_anchored_text_and_author(self) -> None:
        """Read the comment entry and match its identifier to the body reference.
        读取批注入口，并验证标识与正文引用对应。
        """
        # Page into the paragraph that anchors the fixture's single comment.
        # 定位到包含样例唯一批注锚点的正文段落。
        body = self.result("docx-read-blocks", "office.docx", start_block=2, limit=1)
        comments = self.result("docx-list-comments", "office.docx")
        self.assertEqual(comments["comment_count"], 1)
        # Compare the ID across entries so renumbering the source cannot invalidate the test.
        # 跨入口比较 ID，避免源文件重新编号影响测试。
        comment = comments["comments"][0]
        self.assertEqual([comment["id"]], body["blocks"][0]["comment_ids"])
        self.assertEqual(comment["author"], "Reviewer")
        self.assertEqual(comment["text"], "Confirm the figures.")

    def test_powerpoint_uses_presentation_order_and_shape_text(self) -> None:
        """Read reordered and hidden slides, grouped text, and table cells.
        读取重新排序及隐藏的幻灯片、组合文本和表格单元格。
        """
        # The first presented slide is stored as the second ZIP slide member.
        # 放映顺序的第一页存储于 ZIP 中的第二个幻灯片成员。
        first = self.result("pptx-extract-text", "office.pptx", limit=1)
        self.assertEqual(first["slide_count"], 2)
        self.assertEqual(first["next_slide"], 2)
        # The shape tree contains grouped text and a table, while notes live in a separate part.
        # 形状树包含组合文本和表格，演讲者备注位于独立部件。
        slide = first["slides"][0]
        self.assertTrue(slide["hidden"])
        self.assertEqual(slide["text"].splitlines(), ["Grouped content", "Item", "Total"])
        # Resume from the cursor to confirm exact slide membership and ordering.
        # 从游标处继续读取，确认范围和顺序。
        second = self.result("pptx-extract-text", "office.pptx", start_slide=first["next_slide"])
        self.assertEqual(second["slides"][0]["text"], "Original first")
        self.assertFalse(second["slides"][0]["hidden"])
        self.assertIsNone(second["next_slide"])


    def test_powerpoint_rejects_external_slide_relationships(self) -> None:
        """Reject external slide targets at the managed tool boundary.
        在受管工具边界拒绝外部幻灯片目标。
        """
        with TemporaryDirectory() as directory:
            # Change only the slide relationships of the synthetic fixture.
            # 仅修改合成样例中的幻灯片关系。
            path = Path(directory) / "external-slide.pptx"
            relations_name = "ppt/_rels/presentation.xml.rels"
            with ZipFile(FIXTURES / "office.pptx") as source, ZipFile(path, "w") as output:
                for member in source.infolist():
                    # The archive fixture never needs to reach an actual network host.
                    # 归档样例无需访问真实网络主机。
                    content = source.read(member)
                    if member.filename == relations_name:
                        root = ElementTree.fromstring(content)
                        for record in root:
                            if record.attrib["Type"].endswith("/slide"):
                                record.set("TargetMode", "External")
                                record.set("Target", "https://example.invalid/slide.xml")
                        content = ElementTree.tostring(root)
                    output.writestr(member.filename, content)
            # Direct the real debugger to the amended input package.
            # 让实际调试器读取修改后的输入包。
            result = call_tool("pptx-extract-text", {"path": str(path)})
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("internal Office package relationship", result.stdout + result.stderr)


    def test_pdf_continuation_recovers_full_page_text(self) -> None:
        """Read and reassemble all stored text from a long synthetic PDF page.
        读取并重组较长合成 PDF 页面中的全部存储文本。
        """
        with TemporaryDirectory() as directory:
            # The expected text exceeds the per-call budget.
            # 预期文本长度超过单次调用预算。
            path = Path(directory) / "long.pdf"
            expected = "Office page text " * 90
            write_pdf_fixture(path, expected)
            first = self.result("pdf-extract-text", path, start_page=1, end_page=1)["pages"][0]
            self.assertTrue(first["text_truncated"])
            # Resume the exact page at the offset returned by the first call.
            # 使用第一次调用返回的偏移续读同一页。
            second = self.result("pdf-extract-text", path, start_page=1, end_page=1,
                                 text_offset=first["next_text_offset"])["pages"][0]
            self.assertEqual(first["text"] + second["text"], expected)
            self.assertIsNone(second["next_text_offset"])

    def test_sheet_discovery_returns_a_usable_next_index(self) -> None:
        """Use one list page to discover the next hidden worksheet.
        使用一个工作表列表页定位下一张隐藏工作表。
        """
        # The first page contains the data sheet and exposes the hidden sheet's index.
        # 第一页包含数据表并返回下一张隐藏表的索引。
        first = self.result("xlsx-list-sheets", "office.xlsx", limit=1)
        self.assertEqual(first["sheets"][0]["name"], "Data")
        second = self.result("xlsx-list-sheets", "office.xlsx", start_sheet=first["next_sheet"])
        self.assertEqual(second["sheets"], [{"name": "Hidden", "visibility": "hidden"}])
        self.assertIsNone(second["next_sheet"])

    def test_excel_continuation_preserves_unicode_text(self) -> None:
        """Recover the full string value of a cell using its returned continuation.
        使用续读参数恢复单元格的完整字符串值。
        """
        with TemporaryDirectory() as directory:
            # Multibyte characters make byte and character offsets observably different.
            # 多字节字符使字节偏移和字符偏移可以被区分。
            path = Path(directory) / "long.xlsx"
            expected = "审核状态" * 58
            replace_fixture_xml(FIXTURES / "office.xlsx", path, "xl/worksheets/sheet1.xml", "Label", expected)
            first = self.result("xlsx-read-range", path, sheet="Data", range="A1")["cells"][0]
            self.assertTrue(first["value_truncated"])
            second = self.result("xlsx-read-range", path, sheet="Data", range="A1",
                                 text_offset=first["next_value_offset"])["cells"][0]
            self.assertEqual(first["value"] + second["value"], expected)
            self.assertIsNone(second["next_value_offset"])

    def test_word_continuation_recovers_full_paragraph(self) -> None:
        """Resume a long Word block without re-reading the preceding characters.
        续读较长 Word 内容块，避免重复读取前段字符。
        """
        with TemporaryDirectory() as directory:
            # Only the body fixture's first text run changes.
            # 仅替换正文样例的首个文本片段。
            path = Path(directory) / "long.docx"
            expected = "工作文档" * 280
            replace_fixture_xml(FIXTURES / "office.docx", path, "word/document.xml", "Vulcan Work", expected)
            first = self.result("docx-read-blocks", path, limit=1)["blocks"][0]
            self.assertEqual(first["style_id"], "Heading1")
            second = self.result("docx-read-blocks", path, limit=1,
                                 text_offset=first["next_text_offset"])["blocks"][0]
            self.assertEqual(first["text"] + second["text"], expected)
            self.assertIsNone(second["next_text_offset"])

    def test_powerpoint_notes_are_optional_and_can_be_resumed(self) -> None:
        """Keep speaker notes separate and paginate their text independently.
        单独返回演讲者备注，并独立继续读取备注文本。
        """
        with TemporaryDirectory() as directory:
            # The fixture's first presented slide is the one with a stored notes part.
            # 样例按放映顺序的第一页具有已保存的备注部件。
            path = Path(directory) / "notes.pptx"
            expected = "演讲提示" * 270
            replace_fixture_xml(FIXTURES / "office.pptx", path, "ppt/notesSlides/notesSlide1.xml",
                                "Speaker notes are outside slide text", expected)
            plain = self.result("pptx-extract-text", path, limit=1)["slides"][0]
            self.assertNotIn("notes", plain)
            first = self.result("pptx-extract-text", path, limit=1, include_notes=True)["slides"][0]
            self.assertEqual(first["text"], plain["text"])
            self.assertTrue(first["notes"]["text_truncated"])
            second = self.result("pptx-extract-text", path, limit=1, include_notes=True,
                                 notes_offset=first["notes"]["next_text_offset"])["slides"][0]
            self.assertEqual(first["notes"]["text"] + second["notes"]["text"], expected)
            self.assertIsNone(second["notes"]["next_text_offset"])


    def test_pdf_compose_copies_pages_in_requested_order(self) -> None:
        """Write a merged PDF, then read the generated pages and check the originals.
        写入合并 PDF，再读回生成的页面并核对源文件。
        """
        with TemporaryDirectory() as directory:
            # Two different synthetic PDFs make ordering observable.
            # 两份不同的合成 PDF 使页面顺序可以验证。
            folder = Path(directory)
            first = folder / "first.pdf"
            second = folder / "second.pdf"
            output = folder / "combined.pdf"
            write_pdf_fixture(first, "First page")
            write_pdf_fixture(second, "Second page")
            before = first.read_bytes(), second.read_bytes()
            result = self.invoke("pdf-compose", sources=[{"path": str(second)}, {"path": str(first), "start_page": 1}],
                                 output_path=str(output))
            self.assertEqual(result["page_count"], 2)
            pages = self.result("pdf-extract-text", output, start_page=1, end_page=2)["pages"]
            self.assertEqual([page["text"] for page in pages], ["Second page", "First page"])
            self.assertEqual((first.read_bytes(), second.read_bytes()), before)

    def test_office_writers_never_replace_an_existing_output(self) -> None:
        """Verify that an existing output remains byte-for-byte unchanged.
        验证已有输出文件的字节完全不变。
        """
        with TemporaryDirectory() as directory:
            # Occupy the path with a sentinel before calling the real writer.
            # 在调用实际写入工具前用哨兵内容占用输出路径。
            output = Path(directory) / "existing.pdf"
            output.write_bytes(b"keep existing bytes")
            result = call_tool("pdf-compose", {"sources": [{"path": str(FIXTURES / "office.pdf")}],
                                               "output_path": str(output)})
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("already exists", result.stdout + result.stderr)
            self.assertEqual(output.read_bytes(), b"keep existing bytes")
            self.assertEqual(list(output.parent.glob(".vulcan-work-*.tmp")), [])

    def test_excel_writer_distinguishes_literal_text_and_formulas(self) -> None:
        """Create a workbook and read literal values and explicit formulas back.
        新建工作簿并读回字面量值和显式公式。
        """
        with TemporaryDirectory() as directory:
            # A leading equals sign is ordinary text unless supplied through formula.
            # 除非通过 formula 提供，开头的等号仍是普通文本。
            output = Path(directory) / "汇总.xlsx"
            result = self.invoke("xlsx-write-cells", output_path=str(output), sheet="汇总", cells=[
                {"address": "A1", "value": "=1+1"}, {"address": "B1", "formula": "=1+1"},
                {"address": "C1", "value": True},
            ])
            self.assertEqual(result["cells_written"], 3)
            data = self.result("xlsx-read-range", output, sheet="汇总", range="A1:C1")
            literal, formula, boolean = data["cells"]
            self.assertEqual(literal["value"], "=1+1")
            self.assertEqual(formula["formula"], "=1+1")
            self.assertIsNone(formula["cached_value"])
            self.assertIs(boolean["value"], True)

    def test_excel_writer_clears_a_cell_in_a_new_copy(self) -> None:
        """Clear a cell while preserving the source file and other worksheets.
        清除一个单元格，同时保留源文件和其他工作表。
        """
        with TemporaryDirectory() as directory:
            # Use the existing workbook fixture and save to a separate path.
            # 使用现有工作簿样例并保存到独立路径。
            source = FIXTURES / "office.xlsx"
            before = source.read_bytes()
            output = Path(directory) / "edited.xlsx"
            self.invoke("xlsx-write-cells", path=str(source), output_path=str(output), sheet="Data",
                        cells=[{"address": "A1", "value": None}])
            cells = self.result("xlsx-read-range", output, sheet="Data", range="A1:B1")["cells"]
            self.assertIsNone(cells[0]["value"])
            self.assertEqual(cells[1]["value"], 6)
            self.assertEqual(self.result("xlsx-list-sheets", output)["sheet_count"], 2)
            self.assertEqual(source.read_bytes(), before)

    def test_word_writer_creates_and_appends_to_a_copy(self) -> None:
        """Create a Word document and append formatted paragraphs to its copy.
        新建 Word 文档，并向其副本追加带结构的段落。
        """
        with TemporaryDirectory() as directory:
            # Use the public tool to create the initial Word document.
            # 使用公开工具创建初始 Word 文档。
            source = Path(directory) / "new.docx"
            output = source.with_name("edited.docx")
            self.invoke("docx-append-paragraphs", output_path=str(source),
                        paragraphs=[{"text": "进展", "heading_level": 2}])
            before = source.read_bytes()
            self.invoke("docx-append-paragraphs", path=str(source), output_path=str(output),
                        paragraphs=[{"text": "第一行\n第二行\t确认为空格之外的制表符"}])
            blocks = self.result("docx-read-blocks", output)["blocks"]
            self.assertEqual(blocks[0]["style_id"], "Heading2")
            self.assertEqual(blocks[0]["outline_level"], 1)
            self.assertEqual(blocks[1]["text"], "第一行\n第二行\t确认为空格之外的制表符")
            self.assertEqual(source.read_bytes(), before)


    def test_powerpoint_writer_replaces_across_styled_runs(self) -> None:
        """Replace text spanning two runs and preserve unrelated slide content.
        替换跨两个文本片段的文字，并保留无关的幻灯片内容。
        """
        with TemporaryDirectory() as directory:
            # The first slide in presentation order is stored as slide2.xml.
            # 放映顺序的第一页存储在 slide2.xml 中。
            source = Path(directory) / "source.pptx"
            output = source.with_name("updated.pptx")
            replace_fixture_xml(FIXTURES / "office.pptx", source, "ppt/slides/slide2.xml",
                                "Grouped content", 'Grouped </a:t></a:r><a:r><a:rPr b="1"/><a:t>content')
            before = source.read_bytes()
            data = self.invoke("pptx-replace-text", path=str(source), output_path=str(output), slide=1,
                               find="Grouped content", replacement="已审核", expected_matches=1)
            self.assertEqual(data["matches_replaced"], 1)
            first, second = self.result("pptx-extract-text", output, include_notes=True)["slides"]
            self.assertEqual(first["text"].splitlines(), ["已审核", "Item", "Total"])
            self.assertTrue(first["notes"]["text"])
            self.assertEqual(second["text"], "Original first")
            self.assertEqual(source.read_bytes(), before)

    def test_powerpoint_writer_rejects_a_stale_match_count(self) -> None:
        """Reject a changed match count without publishing an output file.
        拒绝有变化的匹配数量，并且不发布输出文件。
        """
        with TemporaryDirectory() as directory:
            # Ask to replace two occurrences where the fixture has exactly one.
            # 样例只有一处匹配，故意请求替换两处。
            output = Path(directory) / "unexpected.pptx"
            result = call_tool("pptx-replace-text", {
                "path": str(FIXTURES / "office.pptx"), "output_path": str(output), "slide": 1,
                "find": "Grouped content", "replacement": "Updated", "expected_matches": 2,
            })
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("found 1", result.stdout + result.stderr)
            self.assertFalse(output.exists())
