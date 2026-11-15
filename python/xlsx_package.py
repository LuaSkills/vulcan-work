"""Update Excel cell XML without rewriting unrelated workbook parts.
更新 Excel 单元格 XML，同时保留无关的工作簿部件。
"""

from collections.abc import Iterator
from xml.dom.minidom import Element
from zipfile import ZipFile

from defusedxml.minidom import parseString
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter, range_boundaries

from python.local_files import read_office_zip
from python.office_documents import (main_part, namespace, related_part, relationship_base,
                                     relationship_target, relationships, root_variant, xml_part)
from python.opc_dom import REL, TYPES, add_dom, nodes, rels_name
from python.write_documents import XML, checked_package, rewrite_package


def direct_value(parent: Element, uri: str, local: str, value: str) -> Element:
    """Append a worksheet element with one text value.
    追加具有一个文本值的工作表元素。
    """
    element = add_dom(parent, uri, local)
    element.appendChild(parent.ownerDocument.createTextNode(value))
    return element


def sheet_cells(sheet_data: Element, uri: str) -> Iterator[Element]:
    """Yield physical cells from direct data rows in worksheet order.
    按工作表顺序逐个返回直接数据行中的实际单元格。
    """
    for row in nodes(sheet_data, uri, "row"):
        yield from nodes(row, uri, "c")


def protect_ranges(sheet: Element, sheet_data: Element, uri: str, addresses: list[str]) -> None:
    """Reject changes in merged non-anchor cells or multi-cell formula regions.
    拒绝修改合并区域非锚点单元格以及跨单元格公式区域。
    """
    # Use ranges declared by worksheet metadata and formula anchors.
    # 使用工作表元数据和公式锚点声明的范围。
    merged = [item.getAttribute("ref") for group in nodes(sheet, uri, "mergeCells")
              for item in nodes(group, uri, "mergeCell")]
    formula_ranges = [formula.getAttribute("ref") for cell in sheet_cells(sheet_data, uri)
                      for formula in nodes(cell, uri, "f")
                      if formula.getAttribute("t") in ("array", "dataTable", "shared")
                      and formula.hasAttribute("ref")]
    regions = [(kind, range_boundaries(ref)) for kind, ranges in (("merged", merged), ("formula", formula_ranges))
               for ref in ranges]
    for address in addresses:
        # Explicit coordinates keep insertion and range checks consistent with the read tools.
        # 显式坐标使插入及范围校验与读取工具保持一致。
        row, column = coordinate_to_tuple(address)
        for kind, (left, top, right, bottom) in regions:
            if left <= column <= right and top <= row <= bottom:
                if kind == "formula" or (column, row) != (left, top):
                    raise ValueError(f"Cannot edit {address} inside a {kind} cell range.")


def ordered_cell(sheet_data: Element, uri: str, address: str) -> Element:
    """Find or insert an explicitly addressed cell without changing other columns.
    查找或插入具有显式地址的单元格，同时保留其他列。
    """
    # The table is stored in row and column order.
    # 数据表按行列顺序存储。
    number, column = coordinate_to_tuple(address)
    row = None
    for existing in nodes(sheet_data, uri, "row"):
        row_number = int(existing.getAttribute("r"))
        if row_number == number:
            row = existing
            break
        if row_number > number:
            row = add_dom(sheet_data, uri, "row", r=number)
            sheet_data.insertBefore(row, existing)
            break
    if row is None:
        row = add_dom(sheet_data, uri, "row", r=number)
    if row.hasAttribute("spans"):
        row.removeAttribute("spans")
    # Position a new cell before the first greater column.
    # 将新单元格放在首个列号更大的单元格之前。
    for cell in nodes(row, uri, "c"):
        existing_column = coordinate_to_tuple(cell.getAttribute("r"))[1]
        if existing_column == column:
            return cell
        if existing_column > column:
            inserted = add_dom(row, uri, "c", r=address)
            row.insertBefore(inserted, cell)
            return inserted
    # An optional row extension list must remain after the worksheet cells.
    # 可选的行扩展列表必须继续位于工作表单元格之后。
    extensions = nodes(row, uri, "extLst")
    inserted = add_dom(row, uri, "c", r=address)
    if extensions:
        row.insertBefore(inserted, extensions[0])
    return inserted


def write_cell(cell: Element, uri: str, value: object, is_formula: bool) -> bool:
    """Write a scalar or formula while retaining cell style and other attributes.
    写入标量或公式，同时保留单元格样式和其他属性。
    """
    # Rich value metadata or a grouped formula requires a group-aware editor.
    # 富值元数据或成组公式需要能够理解整组内容的编辑器。
    if cell.hasAttribute("cm") or cell.hasAttribute("vm") or any(
        item.getAttribute("t") in ("array", "dataTable", "shared") for item in nodes(cell, uri, "f")
    ):
        raise ValueError(f"Complex value or formula metadata at {cell.getAttribute('r')} cannot be edited.")
    replaced_shared = cell.getAttribute("t") == "s"
    if cell.hasAttribute("t"):
        cell.removeAttribute("t")
    for kind in ("f", "v", "is"):
        for existing in nodes(cell, uri, kind):
            cell.removeChild(existing)
    # Excel inline strings avoid allocating or rewriting a shared string table.
    # Excel 内联字符串避免分配或重写共享字符串表。
    if is_formula:
        direct_value(cell, uri, "f", value[1:])
    elif isinstance(value, str):
        cell.setAttribute("t", "inlineStr")
        text = direct_value(add_dom(cell, uri, "is"), uri, "t", value)
        text.setAttributeNS(XML, "xml:space", "preserve")
    elif isinstance(value, bool):
        cell.setAttribute("t", "b")
        direct_value(cell, uri, "v", "1" if value else "0")
    elif value is not None:
        direct_value(cell, uri, "v", str(value))
    # Preserve the schema order when an existing cell carries extension metadata.
    # 当已有单元格包含扩展元数据时，保留 Schema 元素顺序。
    extensions = nodes(cell, uri, "extLst")
    if extensions:
        for kind in ("f", "v", "is"):
            for item in nodes(cell, uri, kind):
                cell.insertBefore(item, extensions[0])
    return replaced_shared


def expand_dimension(sheet: Element, uri: str, addresses: list[str]) -> None:
    """Expand an existing worksheet dimension to cover any added cells.
    扩大已有工作表的维度，以覆盖新增单元格。
    """
    dimensions = nodes(sheet, uri, "dimension")
    if not dimensions:
        return
    # The stored dimension is a cache, so retain its original covered region.
    # 存储的维度属于缓存，因此保留其原有覆盖区域。
    left, top, right, bottom = range_boundaries(dimensions[0].getAttribute("ref"))
    for address in addresses:
        row, column = coordinate_to_tuple(address)
        left, top, right, bottom = min(left, column), min(top, row), max(right, column), max(bottom, row)
    first, last = f"{get_column_letter(left)}{top}", f"{get_column_letter(right)}{bottom}"
    dimensions[0].setAttribute("ref", first if first == last else f"{first}:{last}")


def drop_formula_caches(sheet_data: Element, uri: str) -> bool:
    """Clear stale cached formula results without changing the formulas.
    清除过期公式缓存并保留原公式。
    """
    changed = False
    for cell in sheet_cells(sheet_data, uri):
        if nodes(cell, uri, "f"):
            for cached in nodes(cell, uri, "v"):
                cell.removeChild(cached)
                changed = True
    return changed


def reset_calculation(package: ZipFile, source: str, workbook: Element, uri: str,
                      replacements: dict[str, bytes], omitted: set[str]) -> None:
    """Request recalculation and remove the graph of obsolete formula dependencies.
    请求重新计算，并移除过期的公式依赖图。
    """
    # Preserve the user's calculation mode while requesting a full recalculation.
    # 在请求完整重算的同时保留用户的计算模式。
    options = nodes(workbook, uri, "calcPr")
    if options:
        calculation = options[0]
    else:
        calculation = add_dom(workbook, uri, "calcPr")
        after = {"oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes",
                 "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"}
        sibling = next((node for node in workbook.childNodes
                        if node.nodeType == node.ELEMENT_NODE and node.namespaceURI == uri
                        and node.localName in after), None)
        if sibling is not None:
            workbook.insertBefore(calculation, sibling)
    calculation.setAttribute("fullCalcOnLoad", "1")
    calculation.setAttribute("forceFullCalc", "1")
    chain = related_part(package, source, "calcChain")
    if chain is None:
        return
    # Delete the relationship, part and content type together.
    # 一并删除关系、部件和内容类型。
    relationships_filename = rels_name(source)
    relations = parseString(package.read(relationships_filename), forbid_dtd=True)
    for record in nodes(relations.documentElement, REL, "Relationship"):
        if record.getAttribute("Type").endswith("/calcChain"):
            record.parentNode.removeChild(record)
    types = parseString(package.read("[Content_Types].xml"), forbid_dtd=True)
    for record in nodes(types.documentElement, TYPES, "Override"):
        if record.getAttribute("PartName").removeprefix("/") == chain:
            record.parentNode.removeChild(record)
    replacements[relationships_filename] = relations.toxml(encoding="UTF-8")
    replacements["[Content_Types].xml"] = types.toxml(encoding="UTF-8")
    omitted.update((chain, rels_name(chain)))


def worksheet_parts(package: ZipFile, source: str, workbook: Element, uri: str, relation: str) -> dict[str, str]:
    """Map declared worksheet names to their exact internal parts, excluding chart sheets.
    将声明的工作表名称映射到准确内部部件，排除图表工作表。
    """
    # The relationship type, rather than a filename pattern, identifies worksheets.
    # 由关系类型识别工作表，不依赖文件名模式。
    targets = {record.attrib["Id"]: relationship_target(source, record)
               for record in relationships(package, source) if record.get("Type") == relation + "/worksheet"}
    return {sheet.getAttribute("name"): targets[identity]
            for group in nodes(workbook, uri, "sheets") for sheet in nodes(group, uri, "sheet")
            if (identity := sheet.getAttributeNS(relation, "id")) in targets}


def write_workbook_copy(path: str, output_path: str, sheet_name: str, updates: list[tuple[str, object, bool]]) -> dict:
    """Edit selected cells and preserve every unrelated workbook member.
    编辑选定单元格，并保留所有无关的工作簿成员。
    """
    with checked_package(read_office_zip(path, ".xlsx")) as package:
        # Resolve the workbook vocabulary and physical sheet parts before editing.
        # 编辑前解析工作簿词汇表及实际工作表部件。
        source = main_part(package)
        variant = root_variant(xml_part(package, source), "spreadsheetml", "workbook")
        uri = namespace("spreadsheetml", variant)[1:-1]
        workbook_doc = parseString(package.read(source), forbid_dtd=True)
        workbook = workbook_doc.documentElement
        sheets = worksheet_parts(package, source, workbook, uri, relationship_base(variant))
        if sheet_name not in sheets:
            raise ValueError(f"Unknown worksheet: {sheet_name!r}.")
        # Only changed sheets and workbook calculation metadata are serialized.
        # 只序列化有改动的工作表和工作簿计算元数据。
        replacements = {}
        omitted = set()
        replaced_shared = False
        addresses = [address for address, _, _ in updates]
        for name, part in sheets.items():
            document = parseString(package.read(part), forbid_dtd=True)
            data = nodes(document.documentElement, uri, "sheetData")
            if len(data) != 1:
                raise ValueError(f"Worksheet {name!r} must have one sheetData element.")
            if name == sheet_name:
                protect_ranges(document.documentElement, data[0], uri, addresses)
                for address, value, is_formula in updates:
                    cell = ordered_cell(data[0], uri, address)
                    replaced_shared |= write_cell(cell, uri, value, is_formula)
                expand_dimension(document.documentElement, uri, addresses)
            changed = drop_formula_caches(data[0], uri)
            if changed or name == sheet_name:
                replacements[part] = document.toxml(encoding="UTF-8")
        if replaced_shared:
            # Keep shared string indices intact and remove the obsolete reference count.
            # 保留共享字符串索引，并移除已过时的引用总数。
            part = related_part(package, source, "sharedStrings")
            if part is None:
                raise ValueError("A shared string cell has no corresponding string table.")
            strings = parseString(package.read(part), forbid_dtd=True)
            if strings.documentElement.hasAttribute("count"):
                strings.documentElement.removeAttribute("count")
                replacements[part] = strings.toxml(encoding="UTF-8")
        reset_calculation(package, source, workbook, uri, replacements, omitted)
        replacements[source] = workbook_doc.toxml(encoding="UTF-8")
        return rewrite_package(package, replacements, omitted, output_path, ".xlsx")
