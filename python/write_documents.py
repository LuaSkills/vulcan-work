"""Write targeted Office changes to a new DOCX or PPTX package.
将针对性的 Office 文档改动写入新的 DOCX 或 PPTX 包。
"""

from bisect import bisect_left, bisect_right
from io import BytesIO
from itertools import accumulate, islice
from re import escape, finditer, search, split
from xml.dom.minidom import Document, Element, getDOMImplementation
from zipfile import ZIP_DEFLATED, ZipFile

from defusedxml.minidom import parseString

from python.local_files import read_office_zip
from python.office_documents import (
    main_part, namespace, relationship_base, relationship_target, relationships, root_variant,
    xml_part,
)
from python.output_files import write_new_file

# Namespace constants for newly created Word markup and XML attributes.
# 新建 Word 标记与 XML 属性使用的命名空间常量。
WORD_NS = namespace("wordprocessingml", 0)[1:-1]
XMLNS = "http://www.w3.org/2000/xmlns/"
XML = "http://www.w3.org/XML/1998/namespace"

# Bound edit batches and the size of one paragraph or replacement.
# 限制编辑批量及单段或单项替换的长度。
MAX_PARAGRAPHS = 100
MAX_EDIT_CHARS = 10000
MAX_REPLACEMENTS = 100


def checked_text(text: str) -> str:
    """Reject text that cannot be stored in an Office XML part.
    拒绝无法存储于 Office XML 部件的文本。
    """
    if len(text) > MAX_EDIT_CHARS or search(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]", text):
        raise ValueError("Text is too long or contains invalid XML characters.")
    return text


def checked_package(content: bytes) -> ZipFile:
    """Open a ZIP for a rewrite while rejecting duplicate members and signatures.
    为重写打开 ZIP，并拒绝重复成员和数字签名。
    """
    # A changed signed package would no longer carry a valid signature.
    # 修改后的已签名包将失去有效签名。
    package = ZipFile(BytesIO(content))
    try:
        names = package.namelist()
        if len(names) != len(set(names)) or any(name.lower().startswith("_xmlsignatures/") for name in names):
            raise ValueError("Office packages with duplicate members or digital signatures are unsupported for writing.")
        return package
    except Exception:
        package.close()
        raise


def replace_package_member(package: ZipFile, part: str, data: bytes, output_path: str, extension: str) -> dict:
    """Copy an Office archive to a new file with only the selected member replaced.
    复制 Office 归档到新文件，仅替换指定成员。
    """
    def generate(stream) -> None:
        """Write the replacement bytes and copy every other member as stored.
        写入替换字节，并复制其他成员原有内容。
        """
        with ZipFile(stream, "w") as output:
            for member in package.infolist():
                output.writestr(member, data if member.filename == part else package.read(member))

    return write_new_file(output_path, extension, generate)


def child(document: Document, parent: Element, word: str, name: str) -> Element:
    """Append a WordprocessingML child and return it for further population.
    追加 WordprocessingML 子元素并返回它，以便继续填充。
    """
    # Bind every generated name to the existing document's Word vocabulary.
    # 每个生成的名称都绑定到现有文档的 Word 词汇表。
    element = document.createElementNS(word, f"w:{name}")
    parent.appendChild(element)
    return element


def add_word_paragraph(document: Document, body: Element, word: str, specification: dict) -> None:
    """Insert a paragraph before the terminal Word section properties.
    在 Word 末尾的分节属性之前插入一个段落。
    """
    # A local binding also works when the source uses a different namespace prefix.
    # 即使源文档使用其他命名空间前缀，本地绑定仍然有效。
    text = checked_text(specification["text"])
    heading = specification.get("heading_level")
    if heading is not None and not 1 <= heading <= 9:
        raise ValueError("Word heading levels must be between 1 and 9.")
    paragraph = document.createElementNS(word, "w:p")
    paragraph.setAttributeNS(XMLNS, "xmlns:w", word)
    # The optional outline level gives new headings structure in Word and the read tool.
    # 可选的大纲级别为新标题提供 Word 和读取工具可识别的结构。
    if heading is not None:
        properties = child(document, paragraph, word, "pPr")
        style = child(document, properties, word, "pStyle")
        style.setAttributeNS(word, "w:val", f"Heading{heading}")
        outline = child(document, properties, word, "outlineLvl")
        outline.setAttributeNS(word, "w:val", str(heading - 1))
    # Word represents breaks and tabs as dedicated elements within a run.
    # Word 用文本片段中的独立元素表示换行和制表符。
    run = child(document, paragraph, word, "r")
    if heading is not None:
        child(document, child(document, run, word, "rPr"), word, "b")
    for piece in split(r"(\r\n|\r|\n|\t)", text):
        if piece in ("\n", "\r", "\r\n", "\t"):
            child(document, run, word, "tab" if piece == "\t" else "br")
        elif piece:
            node = child(document, run, word, "t")
            node.setAttributeNS(XML, "xml:space", "preserve")
            node.appendChild(document.createTextNode(piece))
    # Section properties must remain the final element in the main body.
    # 分节属性必须继续作为正文中的最后一个元素。
    section = next((node for node in body.childNodes if node.nodeType == node.ELEMENT_NODE
                    and node.namespaceURI == word and node.localName == "sectPr"), None)
    body.insertBefore(paragraph, section)


def docx_append_paragraphs(args: dict, _context: dict) -> dict:
    """Create a Word document or append ordered paragraphs to an existing copy.
    新建 Word 文档或按顺序向现有文档的副本追加段落。
    """
    # A missing source path is an explicit request to generate a blank document.
    # 省略源路径明确表示新建空白文档。
    paragraphs = args["paragraphs"]
    output_path = args["output_path"]
    source_path = args.get("path")
    if not 1 <= len(paragraphs) <= MAX_PARAGRAPHS:
        raise ValueError(f"Provide between 1 and {MAX_PARAGRAPHS} Word paragraphs.")

    if source_path is None:
        # A minimal Word package needs a main part, content types, and a root relationship.
        # 最小 Word 包需要主部件、内容类型和包根关系。
        document = getDOMImplementation().createDocument(WORD_NS, "w:document", None)
        document.documentElement.setAttributeNS(XMLNS, "xmlns:w", WORD_NS)
        body = child(document, document.documentElement, WORD_NS, "body")
        for paragraph in paragraphs:
            add_word_paragraph(document, body, WORD_NS, paragraph)
        result = write_new_file(output_path, ".docx", lambda stream: create_word_package(stream, document))
    else:
        with checked_package(read_office_zip(source_path, ".docx")) as package:
            # Derive the existing document's vocabulary from its main part.
            # 从现有文档的主部件识别其词汇表。
            part = main_part(package)
            variant = root_variant(xml_part(package, part), "wordprocessingml", "document")
            word = namespace("wordprocessingml", variant)[1:-1]
            document = parseString(package.read(part), forbid_dtd=True)
            bodies = document.documentElement.getElementsByTagNameNS(word, "body")
            if len(bodies) != 1:
                raise ValueError("Word document must have one main body.")
            for paragraph in paragraphs:
                add_word_paragraph(document, bodies[0], word, paragraph)
            result = replace_package_member(package, part, document.toxml(encoding="UTF-8"), output_path, ".docx")
    return {**result, "paragraphs_appended": len(paragraphs)}


def create_word_package(stream, document: Document) -> None:
    """Serialize a minimal DOCX package with a generated main document.
    序列化包含生成的主文档的最小 DOCX 包。
    """
    # File extensions and the main relationship supply the OPC loader's part metadata.
    # 文件扩展名和主关系为 OPC 加载器提供部件元数据。
    types = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    root_relationships = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{relationship_base(0)}/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    with ZipFile(stream, "w", compression=ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", types)
        package.writestr("_rels/.rels", root_relationships)
        package.writestr("word/document.xml", document.toxml(encoding="UTF-8"))


def replace_text_runs(runs: list[Element], before: str, after: str) -> int:
    """Replace exact matches across runs, retaining formatting outside each match.
    跨文本片段替换精确匹配，同时保留每处匹配之外的格式。
    """
    # Concatenated positions map to each original run's ending offset.
    # 拼接后的字符位置对应各原始文本片段的结束偏移。
    values = ["".join(node.data for node in run.childNodes if node.nodeType == node.TEXT_NODE)
              for run in runs]
    ends = list(accumulate(map(len, values)))
    matches = list(islice(finditer(escape(before), "".join(values)), MAX_REPLACEMENTS + 1))
    if len(matches) > MAX_REPLACEMENTS:
        raise ValueError(f"At most {MAX_REPLACEMENTS} matches may be replaced per call.")
    # Apply right to left so each earlier match keeps its original indices.
    # 从右到左替换，确保更靠前的匹配保留原始索引。
    for match in reversed(matches):
        # Match positions refer to the original, pre-edit paragraph.
        # 匹配位置对应编辑前的原始段落。
        first, last = match.span()
        # A boundary at the start belongs to the following run; an ending boundary belongs to the preceding run.
        # 起点恰逢边界时属于后一文本片段，终点恰逢边界时属于前一片段。
        start_run = bisect_right(ends, first)
        end_run = bisect_left(ends, last)
        start_at = first - (ends[start_run - 1] if start_run else 0)
        end_at = last - (ends[end_run - 1] if end_run else 0)
        if start_run == end_run:
            values[start_run] = values[start_run][:start_at] + after + values[start_run][end_at:]
        else:
            values[start_run] = values[start_run][:start_at] + after
            values[start_run + 1:end_run] = [""] * (end_run - start_run - 1)
            values[end_run] = values[end_run][end_at:]
    if matches:
        for run, value in zip(runs, values, strict=True):
            # Store the changed text on the original styled run.
            # 将修改后的文本保存在原有样式片段上。
            while run.firstChild:
                run.removeChild(run.firstChild)
            run.appendChild(run.ownerDocument.createTextNode(value))
    return len(matches)


def replace_slide_paragraph(paragraph: Element, drawing: str, before: str, after: str) -> int:
    """Replace text within paragraph segments, stopping at explicit line breaks.
    在段落分段内替换文本，并在显式换行处断开。
    """
    # A match can cross adjacent styled runs, but cannot cross a break or field.
    # 匹配可以跨相邻的样式片段，但不能跨过换行或字段。
    runs = []
    changed = 0
    for element in paragraph.childNodes:
        if element.nodeType != element.ELEMENT_NODE or element.namespaceURI != drawing:
            continue
        if element.localName in ("br", "fld"):
            changed += replace_text_runs(runs, before, after)
            runs.clear()
        elif element.localName == "r":
            runs.extend(element.getElementsByTagNameNS(drawing, "t"))
    return changed + replace_text_runs(runs, before, after)


def pptx_replace_text(args: dict, _context: dict) -> dict:
    """Replace a known number of exact text matches on one selected slide.
    在一张选定幻灯片上替换数量已知的精确文本匹配。
    """
    # A confirmed match count makes stale or ambiguous changes fail before publication.
    # 使用确认过的匹配数量，在发布前拒绝过时或歧义的修改。
    before = checked_text(args["find"])
    after = checked_text(args["replacement"])
    expected = args["expected_matches"]
    if not before or any(character in before + after for character in ("\r", "\n", "\t")):
        raise ValueError("Provide nonempty find text and keep the replacement on one line without tabs.")
    if not 1 <= expected <= MAX_REPLACEMENTS:
        raise ValueError(f"Expected matches must be between 1 and {MAX_REPLACEMENTS}.")

    with checked_package(read_office_zip(args["path"], ".pptx")) as package:
        # Resolve the slide from the presentation's ordered relationships.
        # 根据演示文稿中有序的关系解析幻灯片。
        source = main_part(package)
        root = xml_part(package, source)
        variant = root_variant(root, "presentationml", "presentation")
        presentation = namespace("presentationml", variant)
        relation = relationship_base(variant)
        slide_ids = root.findall(f"{presentation}sldIdLst/{presentation}sldId")
        slide_number = args["slide"]
        if not 1 <= slide_number <= len(slide_ids):
            raise ValueError(f"Slide {slide_number} is outside the {len(slide_ids)} available slides.")
        # Only a unique internal slide relationship may identify the mutable member.
        # 只使用唯一的内部幻灯片关系确定可修改的成员。
        relation_id = slide_ids[slide_number - 1].attrib["{" + relation + "}id"]
        candidates = [record for record in relationships(package, source)
                      if record.get("Id") == relation_id and record.get("Type") == relation + "/slide"]
        if len(candidates) != 1:
            raise ValueError("Expected one slide relationship for the selected position.")
        part = relationship_target(source, candidates[0])
        document = parseString(package.read(part), forbid_dtd=True)
        # Preserve the original DOM namespace bindings needed by Office extension markup.
        # 保留原始 DOM 命名空间绑定，以支持 Office 扩展标记。
        slide_ns = presentation[1:-1]
        if (document.documentElement.namespaceURI, document.documentElement.localName) != (slide_ns, "sld"):
            raise ValueError("The selected member is not a PowerPoint slide.")
        trees = document.documentElement.getElementsByTagNameNS(slide_ns, "spTree")
        if len(trees) != 1:
            raise ValueError("The selected slide must have one shape tree.")
        drawing = namespace("drawingml", variant)[1:-1]
        # Count only editable, stored text runs in the selected shape tree.
        # 只统计选定形状树中可编辑的已存储文本片段。
        actual = sum(replace_slide_paragraph(paragraph, drawing, before, after)
                     for paragraph in trees[0].getElementsByTagNameNS(drawing, "p"))
        if actual != expected:
            raise ValueError(f"Expected {expected} matching text occurrences, found {actual}.")
        result = replace_package_member(package, part, document.toxml(encoding="UTF-8"), args["output_path"], ".pptx")
    return {**result, "slide": slide_number, "matches_replaced": actual}
