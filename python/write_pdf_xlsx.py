"""Create PDF page collections and write cells into new workbook copies.
创建 PDF 页面集合，并将单元格写入新的工作簿副本。
"""

from contextlib import closing
from io import BytesIO
from re import fullmatch
from warnings import catch_warnings, simplefilter
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from pypdf import PdfWriter

from python.local_files import read_office_zip
from python.office import open_pdf
from python.output_files import write_new_file

# Keep each call bounded regardless of which host validates the schema.
# 无论由哪个宿主校验 Schema，都限制每次调用的规模。
MAX_PDF_SOURCES = 10
MAX_COMPOSE_PAGES = 500
MAX_CELL_UPDATES = 100

# Drawing and extension parts can be lost by an ordinary openpyxl round trip.
# openpyxl 普通往返保存可能丢失绘图与扩展部件。
UNSUPPORTED_XLSX_PARTS = ("xl/drawings/", "xl/activeX/", "xl/ctrlProps/", "xl/embeddings/",
                          "xl/pivot", "xl/slicer", "customXml/", "_xmlsignatures/")


def pdf_compose(args: dict, _context: dict) -> dict:
    """Write a new PDF from the ordered source files or inclusive page ranges.
    按输入文件或包含端点的页码范围生成新的 PDF。
    """
    # The order in sources is the order of pages in the resulting document.
    # sources 的顺序决定输出文档中的页序。
    sources = args["sources"]
    if not 1 <= len(sources) <= MAX_PDF_SOURCES:
        raise ValueError(f"Provide between 1 and {MAX_PDF_SOURCES} PDF sources.")
    with closing(PdfWriter()) as writer:
        for source in sources:
            # First validate this range against the actual source page count.
            # 先根据源 PDF 的实际总页数校验当前范围。
            reader = open_pdf(source["path"])
            page_count = len(reader.pages)
            first = source.get("start_page", 1)
            last = source.get("end_page", page_count)
            if not 1 <= first <= last <= page_count:
                raise ValueError(f"Select an ascending page range inside the PDF's {page_count} pages.")
            if len(writer.pages) + last - first + 1 > MAX_COMPOSE_PAGES:
                raise ValueError(f"Compose at most {MAX_COMPOSE_PAGES} pages per call.")
            # pypdf expects a half-open zero-based tuple; the public range is one-based and inclusive.
            # pypdf 接收零基半开元组，公开参数则是一基且包含端点。
            writer.append(reader, pages=(first - 1, last))
        return {**write_new_file(args["output_path"], ".pdf", writer.write), "page_count": len(writer.pages)}


def workbook_for_write(source_path: str | None) -> Workbook:
    """Open a supported existing workbook, or create a workbook when omitted.
    打开受支持的工作簿，省略路径时创建新工作簿。
    """
    if source_path is None:
        return Workbook()

    # Guard package features known to be unsafe for a rewrite before loading it.
    # 在加载前拒绝已知无法安全往返保存的包结构。
    content = read_office_zip(source_path, ".xlsx")
    with ZipFile(BytesIO(content)) as package:
        if any(member.startswith(UNSUPPORTED_XLSX_PARTS) for member in package.namelist()):
            raise ValueError("This workbook contains drawings, signatures, or advanced parts unsafe to rewrite.")
    # Treat preservation warnings as failures so no incomplete round trip is published.
    # 将内容保留相关警告视为错误，避免发布不完整的往返副本。
    with catch_warnings():
        simplefilter("error", UserWarning)
        return load_workbook(BytesIO(content), keep_links=True, rich_text=True)


def xlsx_write_cells(args: dict, _context: dict) -> dict:
    """Write scalar values or explicit formulas to one worksheet and publish a copy.
    向一张工作表写入标量值或显式公式，并发布一个新副本。
    """
    # A missing input path explicitly requests a new workbook with one named sheet.
    # 省略输入路径表示新建仅含一张指定工作表的工作簿。
    source_path = args.get("path")
    sheet_name = args["sheet"]
    updates = args["cells"]
    if not 1 <= len(sheet_name) <= 31:
        raise ValueError("Excel worksheet names must contain between 1 and 31 characters.")
    if not 1 <= len(updates) <= MAX_CELL_UPDATES:
        raise ValueError(f"Provide between 1 and {MAX_CELL_UPDATES} cell updates.")
    with closing(workbook_for_write(source_path)) as workbook:
        if source_path is None:
            workbook.active.title = sheet_name
        if sheet_name not in (sheet.title for sheet in workbook.worksheets):
            raise ValueError(f"Unknown worksheet: {sheet_name!r}.")
        # Record canonical addresses to reject contradictions and duplicate writes.
        # 记录规范地址，拒绝冲突或重复写入。
        sheet = workbook[sheet_name]
        addresses = set()
        for update in updates:
            # This tool targets one A1 cell per update, within Excel's actual grid.
            # 每项更新定位 Excel 网格范围内的单个 A1 单元格。
            address = update["address"].upper()
            if not fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", address):
                raise ValueError("Each update must specify one A1 cell address.")
            row, column = coordinate_to_tuple(address)
            if row > 1048576 or column > 16384 or address in addresses:
                raise ValueError(f"Invalid or duplicate Excel cell address: {address}.")
            addresses.add(address)
            # Exactly one explicit value or formula makes literal strings safe by default.
            # 必须提供值或显式公式之一，字符串默认按字面量处理。
            has_value = "value" in update
            has_formula = "formula" in update
            if has_value == has_formula:
                raise ValueError(f"Specify exactly one value or formula for {address}.")
            value = update["value"] if has_value else update["formula"]
            if isinstance(value, str) and len(value) > 32767:
                raise ValueError(f"Excel text exceeds the cell limit at {address}.")
            if has_formula and (not isinstance(value, str) or not value.startswith("=") or not 2 <= len(value) <= 8192):
                raise ValueError(f"An explicit formula at {address} must begin with '=' and fit in 8192 characters.")
            # Preserve user-entered strings that happen to look like formulas or error tokens.
            # 保留看起来像公式或错误标记的用户字符串。
            cell = sheet.cell(row, column)
            cell.value = value
            if has_value and isinstance(value, str):
                cell.data_type = "s"
        return {
            **write_new_file(args["output_path"], ".xlsx", workbook.save),
            "sheet": sheet_name,
            "cells_written": len(addresses),
        }
