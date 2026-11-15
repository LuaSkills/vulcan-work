"""Read PDF and XLSX files through fixed managed-runtime handlers.
通过固定的受管运行时处理器读取 PDF 与 XLSX 文件。
"""

from contextlib import closing
from datetime import date, datetime, time, timedelta
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from pypdf import PdfReader

# Limit bytes read from the local file before a parser sees it.
# 在解析前限制从本地文件读取的字节数。
MAX_INPUT_BYTES = 32 * 1024 * 1024
# Limit the XLSX container's advertised expanded size before reading worksheets.
# 在读取工作表前限制 XLSX 容器声明的解压大小。
MAX_XLSX_EXPANDED_BYTES = 128 * 1024 * 1024
# Limit structural results without concealing that additional items exist.
# 限制结构化结果数量，并显式报告还有其他条目。
MAX_LIST_ITEMS = 100
# Limit both page count and retained text per page.
# 同时限制页数和每页保留的文本量。
MAX_PDF_PAGES = 10
# Limit retained text within a single page.
# 限制单页保留的文本量。
MAX_PDF_PAGE_CHARS = 1000
# Limit both cell count and retained text per cell.
# 同时限制单元格数量和每个单元格保留的文本量。
MAX_XLSX_CELLS = 100
# Limit the text retained for an individual cell value or formula.
# 限制单个单元格值或公式保留的文本量。
MAX_XLSX_CELL_CHARS = 160


def read_local_file(raw_path: str, extension: str) -> bytes:
    """Return bounded bytes from the absolute file path with the expected extension.
    从具有指定扩展名的绝对文件路径读取限定大小的字节内容。
    """
    # An absolute path gives the caller the same interpretation in host and debugger.
    # 绝对路径确保宿主和调试器对调用参数的解释一致。
    path = Path(raw_path)
    if not path.is_absolute() or path.suffix.lower() != extension or not path.is_file():
        raise ValueError(f"Expected an absolute path to a local {extension} file.")
    with path.open("rb") as source:
        # Read one byte beyond the limit to detect oversize files without a separate stat.
        # 多读取一个字节以识别超限文件，无需单独获取文件长度。
        content = source.read(MAX_INPUT_BYTES + 1)
    if len(content) > MAX_INPUT_BYTES:
        raise ValueError("Input file exceeds 32 MiB.")
    return content


def open_pdf(path: str) -> PdfReader:
    """Open the PDF at path and return a reader for an unencrypted document.
    打开 path 指向的 PDF，返回未加密文档的读取器。
    """
    # Work from an in-memory snapshot so the source is closed before parsing.
    # 对内存快照进行解析，确保解析前已关闭源文件。
    reader = PdfReader(BytesIO(read_local_file(path, ".pdf")))
    if reader.is_encrypted:
        raise ValueError("Encrypted PDF files are unsupported.")
    return reader


def pdf_inspect(args: dict, _context: dict) -> dict:
    """Return the PDF page count and bounded AcroForm field names from args.path.
    返回 args.path 指定 PDF 的总页数和限定数量的 AcroForm 字段名。
    """
    # Field enumeration uses the interactive form dictionary, independent of page text.
    # 字段枚举使用交互表单字典，与页面文本独立。
    reader = open_pdf(args["path"])
    fields = reader.get_fields() or {}
    return {
        "page_count": len(reader.pages),
        "field_count": len(fields),
        "field_names": list(fields)[:MAX_LIST_ITEMS],
        "fields_truncated": len(fields) > MAX_LIST_ITEMS,
    }


def pdf_extract_text(args: dict, _context: dict) -> dict:
    """Extract bounded text from the inclusive one-based page range in args.
    提取 args 中一基且端点包含的页码范围内的限定长度文本。
    """
    # Validate range size before parsing and document bounds after opening.
    # 解析前验证范围大小，打开后验证文档页数边界。
    start_page, end_page = args["start_page"], args["end_page"]
    if not 1 <= start_page <= end_page < start_page + MAX_PDF_PAGES:
        raise ValueError("Specify an ascending range of one to ten PDF pages.")
    # Resolve the page count once for both validation and the returned metadata.
    # 页数只计算一次，供校验和返回元数据共用。
    reader = open_pdf(args["path"])
    page_count = len(reader.pages)
    if end_page > page_count:
        raise ValueError(f"Requested PDF page {end_page}, but the document has {page_count} pages.")

    # Retain the original numbering and report truncation per returned page.
    # 保留原始页码，并逐页报告是否截断。
    pages = []
    for page_number in range(start_page, end_page + 1):
        # Convert one-based user page numbers to zero-based parser indices.
        # 将用户的一基页码转换为解析器的零基索引。
        page_text = reader.pages[page_number - 1].extract_text()
        pages.append({
            "page": page_number,
            "text": page_text[:MAX_PDF_PAGE_CHARS],
            "text_truncated": len(page_text) > MAX_PDF_PAGE_CHARS,
        })
    return {"page_count": page_count, "pages": pages}


def read_xlsx(path: str) -> bytes:
    """Return one validated XLSX snapshot for the formula and cached-value readers.
    返回一份经过校验的 XLSX 快照，供公式和缓存值读取器共用。
    """
    # Read once so both views refer to exactly the same workbook bytes.
    # 只读取一次，确保两种视图读取同一份工作簿字节。
    content = read_local_file(path, ".xlsx")
    with ZipFile(BytesIO(content)) as archive:
        if sum(member.file_size for member in archive.infolist()) > MAX_XLSX_EXPANDED_BYTES:
            raise ValueError("Expanded XLSX container exceeds 128 MiB.")
    return content


def xlsx_list_sheets(args: dict, _context: dict) -> dict:
    """Return worksheet names and visibility without loading worksheet cell data.
    返回工作表名称与可见性，无需加载工作表中的单元格数据。
    """
    with closing(load_workbook(BytesIO(read_xlsx(args["path"])), read_only=True, keep_links=False)) as workbook:
        # Worksheets exclude chart-only sheets, which have no cell ranges.
        # 工作表列表排除没有单元格范围的图表工作表。
        worksheets = workbook.worksheets
        return {
            "sheet_count": len(worksheets),
            "sheets": [{"name": sheet.title, "visibility": sheet.sheet_state} for sheet in worksheets[:MAX_LIST_ITEMS]],
            "sheets_truncated": len(worksheets) > MAX_LIST_ITEMS,
        }


def cell_value(value: object) -> tuple[object, bool]:
    """Return a JSON-compatible cell value and whether its text was truncated.
    返回兼容 JSON 的单元格值以及文本是否被截断。
    """
    if isinstance(value, (date, datetime, time)):
        return value.isoformat(), False
    if isinstance(value, timedelta):
        return {"seconds": value.total_seconds()}, False
    if isinstance(value, str):
        return value[:MAX_XLSX_CELL_CHARS], len(value) > MAX_XLSX_CELL_CHARS
    return value, False


def serialize_cell(cell: object, cached_cell: object, address: str) -> dict:
    """Combine source and cached cells at address without recalculating a formula.
    合并 address 处的源单元格与缓存单元格，无需重算公式。
    """
    if cell is None or cell.data_type != "f":
        # Missing worksheet rows are represented as blank cells with their original addresses.
        # 缺失的工作表行表示为保留原地址的空白单元格。
        value, truncated = cell_value(cell.value if cell is not None else None)
        return {
            "address": address,
            "kind": "value",
            "type": cell.data_type if value is not None else None,
            "value": value,
            "value_truncated": truncated,
        }

    # Some workbook formula types carry structured metadata rather than a string.
    # 部分工作簿公式类型携带结构化元数据，且没有普通公式字符串。
    formula = cell.value
    formula_kind = "standard"
    if isinstance(formula, ArrayFormula):
        formula_kind = "array"
        formula = formula.text
    elif isinstance(formula, DataTableFormula):
        formula_kind = "data_table"
        formula = None
    # Keep formula display and the stored result independently bounded.
    # 分别限制公式正文与工作簿中保存的结果。
    formula_text, formula_truncated = cell_value(formula)
    cached_value, cached_truncated = cell_value(cached_cell.value)
    return {
        "address": address,
        "kind": "formula",
        "formula_kind": formula_kind,
        "formula": formula_text,
        "formula_truncated": formula_truncated,
        "cached_type": cached_cell.data_type if cached_value is not None else None,
        "cached_value": cached_value,
        "cached_value_truncated": cached_truncated,
    }


def xlsx_read_range(args: dict, _context: dict) -> dict:
    """Read the exact worksheet and rectangular A1 range from args in row order.
    按行读取 args 中指定的工作表与矩形 A1 区域。
    """
    # Enforce Excel's grid dimensions as well as this tool's bounded result size.
    # 同时校验 Excel 网格的维度和工具自身的返回数量上限。
    first_col, first_row, last_col, last_row = range_boundaries(args["range"].upper())
    if not (1 <= first_col <= last_col <= 16384 and 1 <= first_row <= last_row <= 1048576):
        raise ValueError("Cell range must be ascending and inside the Excel grid.")
    if (last_col - first_col + 1) * (last_row - first_row + 1) > MAX_XLSX_CELLS:
        raise ValueError("Select at most 100 XLSX cells per call.")

    # A single snapshot prevents inconsistent formula and cache reads if a file is replaced.
    # 使用同一份快照，避免文件被替换时两种视图不一致。
    content = read_xlsx(args["path"])
    sheet_name = args["sheet"]
    with (
        closing(load_workbook(BytesIO(content), read_only=True, keep_links=False)) as formulas,
        closing(load_workbook(BytesIO(content), read_only=True, keep_links=False, data_only=True)) as cached,
    ):
        if sheet_name not in (sheet.title for sheet in formulas.worksheets):
            raise ValueError(f"Unknown worksheet: {sheet_name!r}.")
        # Do not trust stale worksheet dimensions when the caller provides exact bounds.
        # 调用方提供了精确范围，此时无需依赖可能过时的工作表维度。
        bounds = {"min_row": first_row, "max_row": last_row, "min_col": first_col, "max_col": last_col}
        formula_rows = formulas[sheet_name].iter_rows(**bounds)
        cache_rows = cached[sheet_name].iter_rows(**bounds)
        # openpyxl may stop at the last stored row even when max_row is supplied.
        # 即使传入 max_row，openpyxl 也可能在最后一条已存储行停止迭代。
        row_pairs = zip(formula_rows, cache_rows, strict=True)
        empty_row = (None,) * (last_col - first_col + 1)
        # Pad to the requested rectangle so an empty selection has the same shape as a populated one.
        # 补全到请求的矩形范围，使空范围和有数据的范围具有相同形状。
        cells = []
        for row in range(first_row, last_row + 1):
            # Both readers share the same snapshot and use the same row and column limits.
            # 两个读取器共用同一快照，且使用相同的行列边界。
            formula_row, cache_row = next(row_pairs, (empty_row, empty_row))
            for column, (cell, cached_cell) in enumerate(zip(formula_row, cache_row, strict=True), first_col):
                cells.append(serialize_cell(cell, cached_cell, f"{get_column_letter(column)}{row}"))
    return {"sheet": sheet_name, "range": args["range"].upper(), "cells": cells}
