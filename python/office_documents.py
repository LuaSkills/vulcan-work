"""Read Word and PowerPoint structure through Office package relationships.
通过 Office 包关系读取 Word 和 PowerPoint 结构。
"""

from collections.abc import Iterator
from io import BytesIO
from posixpath import basename, dirname, join, normpath
from urllib.parse import unquote, urlsplit
from xml.etree.ElementTree import Element
from zipfile import ZipFile

from defusedxml.ElementTree import fromstring

from python.local_files import read_office_zip
from python.text_pages import document_text, select_items, single_item_offset

# OPC uses a single namespace for package and part relationships.
# OPC 对包级和部件级关系文件使用同一命名空间。
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
# Pair the Transitional and Strict Office namespace roots by index.
# 按索引对应过渡版与严格版的 Office 命名空间根。
OFFICE = ("http://schemas.openxmlformats.org", "http://purl.oclc.org/ooxml")
# Strict removes the 2006 segment used in Transitional Office vocabularies.
# 严格版省略过渡版 Office 词汇表中的 2006 路径段。
YEAR = ("/2006", "")
# Cap the number of items returned for each content type.
# 分别限制各类内容在单次调用中返回的条目数。
MAX_WORD_ITEMS = 20
MAX_SLIDES = 10


def namespace(vocabulary: str, variant: int) -> str:
    """Return the expanded ElementTree namespace for the selected Office vocabulary.
    返回所选 Office 词汇表的 ElementTree 展开命名空间。
    """
    return "{" + OFFICE[variant] + "/" + vocabulary + YEAR[variant] + "/main}"


def relationship_base(variant: int) -> str:
    """Return the Office relationship type prefix and attribute namespace.
    返回 Office 关系类型前缀及属性命名空间。
    """
    return OFFICE[variant] + "/officeDocument" + YEAR[variant] + "/relationships"


def xml_part(package: ZipFile, name: str) -> Element:
    """Return a package member parsed as XML with DTDs and entities disabled.
    将包成员解析为 XML 并返回，禁用 DTD 与实体。
    """
    return fromstring(package.read(name), forbid_dtd=True)


def root_variant(root: Element, vocabulary: str, name: str) -> int:
    """Return the recognized OOXML variant for the expected document root.
    返回预期文档根节点所属的已知 OOXML 变体。
    """
    for variant in range(len(OFFICE)):
        if root.tag == namespace(vocabulary, variant) + name:
            return variant
    raise ValueError(f"Unsupported Office XML root for {name}.")


def relationships(package: ZipFile, source: str) -> list[Element]:
    """Return source part relationships, or package relationships for an empty source.
    返回源部件的关系，源为空时返回包级关系。
    """
    # OPC derives the relationships path from the source member.
    # OPC 根据源部件路径计算关系文件路径。
    name = join(dirname(source), "_rels", basename(source) + ".rels")
    try:
        root = xml_part(package, name)
    except KeyError:
        return []
    if root.tag != PKG_REL + "Relationships":
        raise ValueError(f"Invalid Office relationship part: {name}.")
    return root.findall(PKG_REL + "Relationship")


def relationship_target(source: str, record: Element) -> str:
    """Resolve an internal relationship target within the input ZIP only.
    仅在输入 ZIP 内解析内部关系目标。
    """
    # Resolve part URIs inside the archive without touching the filesystem or network.
    # 在归档内解析部件 URI，不访问文件系统或网络。
    uri = urlsplit(record.attrib["Target"])
    path = unquote(uri.path, errors="strict")
    if (record.get("TargetMode", "Internal") != "Internal" or uri.scheme or uri.netloc
            or uri.query or uri.fragment or not path or "\\" in path or "\x00" in path):
        raise ValueError("Expected an internal Office package relationship target.")
    # A leading slash anchors the member at the ZIP root.
    # 前导斜杠将成员定位到 ZIP 根目录。
    target = normpath(path[1:] if path.startswith("/") else join(dirname(source), path))
    if target.startswith(("../", "/")) or target in (".", ".."):
        raise ValueError("Office relationship target escapes the package.")
    return target


def related_part(package: ZipFile, source: str, kind: str) -> str | None:
    """Return the unique target of the optional Office relationship of kind.
    返回 kind 指定的可选 Office 关系的唯一目标。
    """
    # External links and media relations do not match the requested Office type.
    # 外部链接和媒体关系不会匹配所请求的 Office 类型。
    types = {relationship_base(variant) + "/" + kind for variant in range(len(OFFICE))}
    matches = [record for record in relationships(package, source) if record.get("Type") in types]
    if len(matches) > 1:
        raise ValueError(f"Expected one Office relationship of type {kind}.")
    return relationship_target(source, matches[0]) if matches else None


def main_part(package: ZipFile) -> str:
    """Return the main document member declared at the package root.
    返回包根关系中声明的主文档成员。
    """
    # The relationship remains valid even when the main XML member is renamed.
    # 即使主 XML 成员重命名，仍以此关系定位。
    part = related_part(package, "", "officeDocument")
    if part is None:
        raise ValueError("Office package has no main document relationship.")
    return part


def word_blocks(container: Element, word: str) -> Iterator[Element]:
    """Yield paragraphs and tables, including blocks inside content controls.
    逐个返回段落和表格，包括内容控件里的块。
    """
    for child in container:
        if child.tag in (word + "p", word + "tbl"):
            yield child
        elif child.tag in (word + "sdt", word + "sdtContent", word + "customXml"):
            yield from word_blocks(child, word)


def paragraph_text(paragraph: Element, text_tag: str, separators: dict[str, str]) -> str:
    """Read text runs and the supplied separator elements in document order.
    按文档顺序读取文本片段和指定的分隔元素。
    """
    return "".join(
        (element.text or "") if element.tag == text_tag else separators[element.tag]
        for element in paragraph.iter()
        if element.tag == text_tag or element.tag in separators
    )


def word_block_text(block: Element, word: str) -> str:
    """Read Word text, using tabs between table cells and newlines between rows.
    读取 Word 文本，表格列间使用制表符，行间使用换行符。
    """
    if block.tag == word + "p":
        return paragraph_text(block, word + "t", {word + "tab": "\t", word + "br": "\n", word + "cr": "\n"})
    # Nested paragraphs and tables retain their enclosing cell positions.
    # 内部段落与嵌套表格保留所在单元格的位置。
    return "\n".join(
        "\t".join(
            "\n".join(word_block_text(child, word) for child in word_blocks(cell, word))
            for cell in row.findall(word + "tc")
        )
        for row in block.findall(word + "tr")
    )


def docx_read_blocks(args: dict, _context: dict) -> dict:
    """Return main-body blocks with their positions, styles, and comment references.
    返回正文块与各自的位置、样式和批注引用。
    """
    with ZipFile(BytesIO(read_office_zip(args["path"], ".docx"))) as package:
        # Resolve the exact main part declared in the archive.
        # 读取归档中声明的主文档部件。
        root = xml_part(package, main_part(package))
    # The root determines the Word vocabulary and the body is required.
    # 根据根节点确定 Word 词汇表，正文为必需节点。
    word = namespace("wordprocessingml", root_variant(root, "wordprocessingml", "document"))
    body = root.find(word + "body")
    if body is None:
        raise ValueError("Word document has no main body.")
    # Only paragraphs and tables contribute to one-based block positions.
    # 仅段落和表格计入一基块位置。
    blocks = list(word_blocks(body, word))
    start = args.get("start_block", 1)
    selected, next_block = select_items(blocks, start, args.get("limit", MAX_WORD_ITEMS), MAX_WORD_ITEMS)
    # Resume only a single, unambiguous block.
    # 仅在明确选中一个块时续读。
    offset = single_item_offset(args.get("text_offset", 0), len(selected))
    # Preserve source order and return direct paragraph properties without computing layout.
    # 保留源顺序并返回直接段落属性，无需计算版式。
    output = []
    for number, block in enumerate(selected, start):
        # Deduplicate repeated comment anchors and references while retaining their order.
        # 对重复的批注锚点和引用去重并保留原顺序。
        comment_ids = list(dict.fromkeys(
            element.attrib[word + "id"] for element in block.iter()
            if element.tag in (word + "commentRangeStart", word + "commentReference")
        ))
        item = {"block": number, "kind": "table", "comment_ids": comment_ids,
                **document_text(word_block_text(block, word), offset)}
        if block.tag == word + "p":
            # Direct style and outline properties can be absent in otherwise valid paragraphs.
            # 合法段落可以省略直接样式与大纲属性。
            style = block.find(f"{word}pPr/{word}pStyle")
            outline = block.find(f"{word}pPr/{word}outlineLvl")
            item.update({
                "kind": "paragraph",
                "style_id": style.get(word + "val") if style is not None else None,
                "outline_level": int(outline.attrib[word + "val"]) if outline is not None else None,
            })
        output.append(item)
    return {"block_count": len(blocks), "next_block": next_block, "blocks": output}


def docx_list_comments(args: dict, _context: dict) -> dict:
    """Return embedded Word comments with stable IDs and their author metadata.
    返回 Word 内嵌批注、稳定 ID 和作者元数据。
    """
    with ZipFile(BytesIO(read_office_zip(args["path"], ".docx"))) as package:
        # The main Word part owns the optional comments relationship.
        # Word 主部件拥有可选的批注关系。
        document = main_part(package)
        variant = root_variant(xml_part(package, document), "wordprocessingml", "document")
        part = related_part(package, document, "comments")
        root = xml_part(package, part) if part is not None else None
    # A missing relationship represents an empty collection of comments.
    # 缺少批注关系表示批注集合为空。
    word = namespace("wordprocessingml", variant)
    if root is not None and root.tag != word + "comments":
        raise ValueError("Unsupported Office XML root for comments.")
    # Keep the comment order stored in Word and paginate independently of body blocks.
    # 保留 Word 存储的批注顺序，独立于正文块分页。
    comments = root.findall(word + "comment") if root is not None else []
    selected, next_comment = select_items(
        comments, args.get("start_comment", 1), args.get("limit", MAX_WORD_ITEMS), MAX_WORD_ITEMS,
    )
    # Comment text uses the same character continuation as main-body blocks.
    # 批注文本与正文块使用相同的字符续读规则。
    offset = single_item_offset(args.get("text_offset", 0), len(selected))
    return {
        "comment_count": len(comments),
        "next_comment": next_comment,
        "comments": [{
            "id": comment.attrib[word + "id"],
            "author": comment.get(word + "author"),
            "date": comment.get(word + "date"),
            **document_text(
                "\n".join(word_block_text(block, word) for block in word_blocks(comment, word)), offset,
            ),
        } for comment in selected],
    }



def slide_shapes(slide: Element, presentation: str) -> Element:
    """Return the slide or notes slide's stored shape tree.
    返回幻灯片或备注页存储的形状树。
    """
    # The shape tree holds text boxes, groups, table cells, and notes placeholders.
    # 形状树包含文本框、组合、表格单元格和备注占位符。
    shapes = slide.find(f"{presentation}cSld/{presentation}spTree")
    if shapes is None:
        raise ValueError("PowerPoint slide has no shape tree.")
    return shapes


def drawing_text(shape: Element, drawing: str) -> str:
    """Return nonempty DrawingML paragraphs from a shape or group of shapes.
    返回形状或形状集合中的非空 DrawingML 段落。
    """
    # Preserve explicit line breaks inside paragraphs without emitting empty placeholders.
    # 保留段落内部的显式换行，并忽略空占位符。
    separators = {drawing + "br": "\n"}
    paragraphs = (
        paragraph_text(paragraph, drawing + "t", separators)
        for paragraph in shape.iter(drawing + "p")
    )
    return "\n".join(text for text in paragraphs if text)


def speaker_notes(package: ZipFile, source: str, presentation: str, drawing: str) -> str:
    """Return speaker text from the slide's optional notes body placeholders.
    读取幻灯片可选备注页正文占位符中的演讲者文本。
    """
    # Notes use the current slide's own relation and may be absent.
    # 备注使用当前幻灯片的独立关系，也可以不存在。
    part = related_part(package, source, "notesSlide")
    if part is None:
        return ""
    # Only body placeholders represent speaker notes; omit headers and slide numbers.
    # 只有正文占位符对应演讲者备注，省略页眉和幻灯片编号。
    notes = xml_part(package, part)
    if notes.tag != presentation + "notes":
        raise ValueError("Invalid PowerPoint notes part.")
    # The placeholder type belongs to the shape's non-visual properties.
    # 占位符类型位于形状的非视觉属性中。
    placeholder_path = f"{presentation}nvSpPr/{presentation}nvPr/{presentation}ph"
    texts = []
    for shape in slide_shapes(notes, presentation).findall(presentation + "sp"):
        placeholder = shape.find(placeholder_path)
        if placeholder is not None and placeholder.get("type") == "body":
            texts.append(drawing_text(shape, drawing))
    return "\n".join(text for text in texts if text)


def pptx_extract_text(args: dict, _context: dict) -> dict:
    """Return text and visibility for a bounded selection of slides in presentation order.
    按演示文稿顺序返回限定范围内的幻灯片文本及可见状态。
    """
    with ZipFile(BytesIO(read_office_zip(args["path"], ".pptx"))) as package:
        # The stored slide IDs, rather than the ZIP member names, set the presentation order.
        # 以存储的幻灯片 ID 确定放映顺序，不依赖 ZIP 成员名称。
        source = main_part(package)
        root = xml_part(package, source)
        variant = root_variant(root, "presentationml", "presentation")
        presentation = namespace("presentationml", variant)
        drawing = namespace("drawingml", variant)
        slides = root.findall(f"{presentation}sldIdLst/{presentation}sldId")
        start = args.get("start_slide", 1)
        selected, next_slide = select_items(slides, start, args.get("limit", MAX_SLIDES), MAX_SLIDES)
        # Slide text and speaker notes have independent continuation offsets.
        # 幻灯片正文和演讲者备注使用独立续读偏移。
        offset = single_item_offset(args.get("text_offset", 0), len(selected))
        include_notes = args.get("include_notes", False)
        notes_offset = single_item_offset(args.get("notes_offset", 0), len(selected))
        if notes_offset and not include_notes:
            raise ValueError("Enable include_notes to resume speaker notes.")
        # Only relationships of the slide type participate in resolution.
        # 解析时只接受幻灯片类型的关系。
        relationship = relationship_base(variant)
        relations = {
            record.attrib["Id"]: record for record in relationships(package, source)
            if record.get("Type") == relationship + "/slide"
        }
        attribute = "{" + relationship + "}id"
        # Preserve one-based presentation positions, including hidden slides.
        # 保留一基演示位置，包括隐藏幻灯片。
        output = []
        for number, slide_id in enumerate(selected, start):
            # A missing slide relationship cannot be replaced by a guessed member name.
            # 幻灯片关系缺失时不能猜测成员名称来替代。
            record = relations.get(slide_id.attrib[attribute])
            if record is None:
                raise ValueError(f"PowerPoint slide {number} has no slide relationship.")
            # Keep the slide member path to resolve its own notes relationship.
            # 保留幻灯片成员路径，用于解析该幻灯片自身的备注关系。
            slide_part = relationship_target(source, record)
            slide = xml_part(package, slide_part)
            if slide.tag != presentation + "sld":
                raise ValueError(f"Invalid XML part for PowerPoint slide {number}.")
            item = {
                "slide": number,
                "hidden": slide.get("show") in ("0", "false"),
                **document_text(drawing_text(slide_shapes(slide, presentation), drawing), offset),
            }
            if include_notes:
                item["notes"] = document_text(speaker_notes(package, slide_part, presentation, drawing), notes_offset)
            output.append(item)
    return {"slide_count": len(slides), "next_slide": next_slide, "slides": output}
