"""Manipulate bounded Office package relationships and XML elements.
操作大小受限的 Office 包关系和 XML 元素。
"""

from posixpath import basename, dirname, join
from zipfile import ZipFile

from python.office_documents import relationship_base, relationship_target, relationships

# OPC package metadata has stable namespaces across Office variants.
# OPC 包元数据在 Office 变体中使用稳定命名空间。
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"


def nodes(parent, uri: str, local: str) -> list:
    """Return direct elements of a given namespace and local name.
    返回指定命名空间和本地名称的直接子元素。
    """
    return [node for node in parent.childNodes if node.nodeType == node.ELEMENT_NODE
            and node.namespaceURI == uri and node.localName == local]


def add_dom(parent, uri: str, tag: str, **attributes):
    """Append an element using the parent's existing namespace prefix.
    使用父级已有的命名空间前缀追加一个元素。
    """
    # The inherited prefix binds to the same vocabulary as the parent.
    # 继承的前缀绑定到与父级相同的词汇表。
    name = f"{parent.prefix}:{tag}" if parent.prefix else tag
    node = parent.ownerDocument.createElementNS(uri, name)
    for key, value in attributes.items():
        node.setAttribute(key, str(value))
    parent.appendChild(node)
    return node


def rels_name(source: str) -> str:
    """Return the OPC relationship member for a given part.
    返回指定部件对应的 OPC 关系成员。
    """
    return join(dirname(source), "_rels", basename(source) + ".rels")


def target_for_id(package: ZipFile, source: str, relation_id: str, kind: str) -> str:
    """Resolve one typed, internal relationship by its stable identifier.
    按稳定标识解析一个指定类型的内部关系。
    """
    # The type suffix matches either supported Office vocabulary variant.
    # 类型后缀可匹配两种受支持的 Office 词汇表变体。
    kinds = {relationship_base(variant) + "/" + kind for variant in (0, 1)}
    matches = [record for record in relationships(package, source)
               if record.get("Id") == relation_id and record.get("Type") in kinds]
    if len(matches) != 1:
        raise ValueError(f"Expected one {kind} relationship for {relation_id!r}.")
    return relationship_target(source, matches[0])
