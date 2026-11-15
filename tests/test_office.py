"""Exercise the actual LuaSkills entries against small synthetic office files.
使用小型合成办公文件测试实际的 LuaSkills 入口。
"""

import json
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest

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


class OfficeEntryTests(unittest.TestCase):
    """Validate manifest-driven calls and their user-visible output contracts.
    验证清单驱动的调用及其面向用户的输出约定。
    """

    def result(self, entry: str, filename: str, **arguments: object) -> dict:
        """Return the decoded user-visible result for a fixture and operation.
        返回指定样例与操作的用户可见结果。
        """
        # The debugger may print bootstrap status before its JSON envelope.
        # 调试器可能在 JSON 外层结构之前打印启动状态。
        finished = call_tool(entry, {"path": str(FIXTURES / filename), **arguments})
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
