"""Create, append, and remove slides while preserving unrelated package parts.
创建、追加和删除幻灯片，同时保留无关的包部件。
"""

from pathlib import Path
from posixpath import dirname, join, relpath
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

from defusedxml.minidom import parseString

from python.local_files import read_office_zip
from python.office_documents import main_part, namespace, relationship_base, relationship_target, relationships, root_variant, xml_part
from python.office_render import convert_with_libreoffice, render_workspace
from python.output_files import write_new_file
from python.opc_dom import REL, TYPES, add_dom, nodes, rels_name, target_for_id
from python.write_documents import XMLNS, checked_package, checked_text, rewrite_package

# Use the exact OPC metadata namespaces when adding or deleting parts.
# 添加或删除部件时使用准确的 OPC 元数据命名空间。
SLIDE_CONTENT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
MAX_SLIDE_EDITS = 30


def slide_content(items: list) -> list[tuple[str, list[str]]]:
    """Validate a batch of titles and body paragraphs before creating slides.
    在创建幻灯片前校验批量标题和正文段落。
    """
    if not 1 <= len(items) <= MAX_SLIDE_EDITS:
        raise ValueError(f"Provide between 1 and {MAX_SLIDE_EDITS} slides.")
    # Validate XML characters once for both creation and append paths.
    # 为新建和追加路径统一校验 XML 字符。
    result = []
    for item in items:
        title = checked_text(item["title"])
        paragraphs = [checked_text(text) for text in item.get("body", [])]
        if len(paragraphs) > 12:
            raise ValueError("Each new slide accepts up to 12 body paragraphs.")
        result.append((title, paragraphs))
    return result


def create_odp(path: Path, slides: list[tuple[str, list[str]]]) -> None:
    """Create a minimal, self-authored OpenDocument presentation for Impress.
    为 Impress 创建一个最小的自行生成 OpenDocument 演示文稿。
    """
    # OpenDocument uses a master page for slide geometry and per-paragraph styles for text.
    # OpenDocument 使用母版页定义幻灯片几何尺寸，并用段落样式定义文本。
    common = ('xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
              'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
              'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2"')
    styles = (f'<office:document-styles {common}><office:automatic-styles>'
              '<style:page-layout style:name="Wide"><style:page-layout-properties fo:page-width="28cm" '
              'fo:page-height="15.75cm" style:print-orientation="landscape"/></style:page-layout>'
              '</office:automatic-styles><office:master-styles>'
              '<style:master-page style:name="Default" style:page-layout-name="Wide"/>'
              '</office:master-styles></office:document-styles>')
    pages = []
    for number, (title, body) in enumerate(slides, 1):
        # Keep title and body in independent text frames with room for ordinary slide copy.
        # 将标题与正文放在独立文本框内，为常用幻灯片文案保留空间。
        body_xml = "".join(f'<text:p text:style-name="Body">{escape(line)}</text:p>' for line in body)
        pages.append(f'<draw:page draw:name="Slide {number}" draw:master-page-name="Default">'
                     '<draw:frame svg:x="1.4cm" svg:y="1cm" svg:width="25.2cm" svg:height="2.3cm">'
                     f'<draw:text-box><text:p text:style-name="Title">{escape(title)}</text:p></draw:text-box></draw:frame>'
                     '<draw:frame svg:x="1.4cm" svg:y="4cm" svg:width="25.2cm" svg:height="10cm">'
                     f'<draw:text-box>{body_xml}</draw:text-box></draw:frame></draw:page>')
    content = (f'<office:document-content {common} xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
               'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
               'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">'
               '<office:automatic-styles><style:style style:name="Title" style:family="paragraph">'
               '<style:text-properties fo:font-size="28pt" fo:font-weight="bold"/></style:style>'
               '<style:style style:name="Body" style:family="paragraph">'
               '<style:text-properties fo:font-size="18pt"/></style:style></office:automatic-styles>'
               '<office:body><office:presentation>' + "".join(pages) + '</office:presentation></office:body></office:document-content>')
    media = "application/vnd.oasis.opendocument.presentation"
    manifest = ('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
                f'<manifest:file-entry manifest:full-path="/" manifest:media-type="{media}"/>'
                '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
                '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>')
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as package:
        package.writestr("mimetype", media, compress_type=ZIP_STORED)
        package.writestr("content.xml", content)
        package.writestr("styles.xml", styles)
        package.writestr("META-INF/manifest.xml", manifest)


def pptx_create(args: dict, _context: dict) -> dict:
    """Create a new title-and-body presentation through LibreOffice Impress.
    通过 LibreOffice Impress 创建带标题与正文的演示文稿。
    """
    slides = slide_content(args["slides"])
    with render_workspace(args.get("_render_temp")) as folder:
        # Conversion uses a private ODP and profile instead of a bundled Office template.
        # 转换使用私有 ODP 和配置，无需随包附带办公模板。
        source = folder / "slides.odp"
        create_odp(source, slides)
        converted = convert_with_libreoffice(source, folder, "pptx:Impress MS PowerPoint 2007 XML", args.get("_soffice"))
        content = read_office_zip(str(converted), ".pptx")
    result = write_new_file(args["output_path"], ".pptx", lambda stream: stream.write(content))
    return {**result, "slide_count": len(slides)}


def presentation_nodes(package: ZipFile):
    """Return the declared presentation, its relationships, and content types.
    返回声明的演示文稿、其关系以及内容类型。
    """
    # Retain original namespace prefixes for markup compatibility in the copied presentation.
    # 保留演示文稿原有命名空间前缀，维持标记兼容性。
    source = main_part(package)
    variant = root_variant(xml_part(package, source), "presentationml", "presentation")
    presentation = namespace("presentationml", variant)[1:-1]
    document = parseString(package.read(source), forbid_dtd=True)
    relation = parseString(package.read(rels_name(source)), forbid_dtd=True)
    types = parseString(package.read("[Content_Types].xml"), forbid_dtd=True)
    return source, variant, presentation, document, relation, types


def blank_slide_xml(variant: int, width: int, height: int, title: str, body: list[str]) -> bytes:
    """Build a slide part with a title and ordinary DrawingML text boxes.
    创建具有标题和常规 DrawingML 文本框的幻灯片部件。
    """
    # Keep the generated shapes within the actual presentation dimensions.
    # 将生成形状限制在当前演示文稿的实际尺寸之内。
    drawing = namespace("drawingml", variant)[1:-1]
    presentation = namespace("presentationml", variant)[1:-1]
    left, content_width = width // 16, width * 7 // 8

    def box(identifier: int, label: str, top: int, box_height: int, lines: list[str], font_size: int) -> str:
        """Serialize a horizontally inset text box.
        序列化一个带水平边距的文本框。
        """
        # A fixed font size is recorded in editable runs rather than rasterized.
        # 固定字号存储于可编辑的文本片段中。
        text = "".join(f'<a:p><a:r><a:rPr sz="{font_size}"/><a:t>{escape(line)}</a:t></a:r></a:p>'
                       for line in lines) or "<a:p/>"
        return (f'<p:sp><p:nvSpPr><p:cNvPr id="{identifier}" name="{label}"/><p:cNvSpPr txBox="1"/>'
                '<p:nvPr/></p:nvSpPr><p:spPr>'
                f'<a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{content_width}" cy="{box_height}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                f'<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>{text}</p:txBody></p:sp>')

    shapes = (box(2, "Title", height // 16, height // 6, [title], 2800)
              + box(3, "Body", height // 4, height * 2 // 3, body, 1800))
    return (f'<p:sld xmlns:p="{presentation}" xmlns:a="{drawing}"><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
            f'{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>').encode("utf-8")


def primary_layout(package: ZipFile, source: str, presentation: str, relation: str, root, slides: list) -> str:
    """Choose the first current slide's layout or the first master layout of an empty deck.
    选择首张幻灯片的布局，空演示文稿则选首个母版布局。
    """
    if slides:
        first = target_for_id(package, source, slides[0].getAttributeNS(relation, "id"), "slide")
        layout = [record for record in relationships(package, first)
                  if record.get("Type") == relation + "/slideLayout"]
        if len(layout) != 1:
            raise ValueError("The first slide must have one slide layout.")
        return relationship_target(first, layout[0])

    # Empty decks still declare their masters and ordered layout lists.
    # 空演示文稿仍声明母版及有序布局列表。
    masters = nodes(root, presentation, "sldMasterIdLst")
    if not masters or not nodes(masters[0], presentation, "sldMasterId"):
        raise ValueError("The presentation has no slide master for new slides.")
    master = nodes(masters[0], presentation, "sldMasterId")[0]
    master_part = target_for_id(package, source, master.getAttributeNS(relation, "id"), "slideMaster")
    master_xml = xml_part(package, master_part)
    layouts = master_xml.findall(f"{{{presentation}}}sldLayoutIdLst/{{{presentation}}}sldLayoutId")
    if not layouts:
        raise ValueError("The slide master has no usable slide layout.")
    return target_for_id(package, master_part, layouts[0].attrib["{" + relation + "}id"], "slideLayout")


def pptx_append_slides(args: dict, _context: dict) -> dict:
    """Append editable title-and-body slides to a copy of an existing presentation.
    将可编辑的标题和正文幻灯片追加到现有演示文稿的副本。
    """
    additions = slide_content(args["slides"])
    with checked_package(read_office_zip(args["path"], ".pptx")) as package:
        # Obtain exact part names instead of assuming the file follows a default folder layout.
        # 获取实际部件名称，避免假定文件使用默认目录结构。
        source, variant, presentation, document, relation_doc, types = presentation_nodes(package)
        relation = relationship_base(variant)
        root = document.documentElement
        lists = nodes(root, presentation, "sldIdLst")
        sizes = nodes(root, presentation, "sldSz")
        if len(lists) != 1 or len(sizes) != 1:
            raise ValueError("Expected one slide list and one slide size.")
        slides = nodes(lists[0], presentation, "sldId")
        layout = primary_layout(package, source, presentation, relation, root, slides)
        width, height = (int(sizes[0].getAttribute(key)) for key in ("cx", "cy"))
        if min(width, height) <= 0:
            raise ValueError("The presentation slide size must be positive.")
        # Allocate distinct IDs in both the presentation list and its relationship table.
        # 在演示文稿列表和关系表中分别分配唯一 ID。
        next_id = max([255, *(int(slide.getAttribute("id")) for slide in slides)]) + 1
        used = {record.getAttribute("Id") for record in nodes(relation_doc.documentElement, REL, "Relationship")}
        archive_names = set(package.namelist())
        replacements = {}
        for title, body in additions:
            # The package member name and relationship ID use independent authoritative collision checks.
            # 包成员名和关系 ID 分别按实际归档检查冲突。
            while True:
                part = join(dirname(source), "slides", f"vulcan-work-{next_id}.xml")
                relation_id = f"rIdVulcanWork{next_id}"
                if part not in archive_names and rels_name(part) not in archive_names and relation_id not in used:
                    break
                next_id += 1
            slide_id = add_dom(lists[0], presentation, "sldId", id=next_id)
            slide_id.setAttributeNS(XMLNS, "xmlns:r", relation)
            slide_id.setAttributeNS(relation, "r:id", relation_id)
            add_dom(relation_doc.documentElement, REL, "Relationship", Id=relation_id,
                    Type=relation + "/slide", Target=relpath(part, dirname(source)))
            add_dom(types.documentElement, TYPES, "Override", PartName="/" + part, ContentType=SLIDE_CONTENT)
            replacements[part] = blank_slide_xml(variant, width, height, title, body)
            relative_layout = escape(relpath(layout, dirname(part)), {'"': "&quot;"})
            replacements[rels_name(part)] = (
                f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{relation}/slideLayout" '
                f'Target="{relative_layout}"/></Relationships>').encode("utf-8")
            next_id += 1
        replacements.update({
            source: document.toxml(encoding="UTF-8"),
            rels_name(source): relation_doc.toxml(encoding="UTF-8"),
            "[Content_Types].xml": types.toxml(encoding="UTF-8"),
        })
        result = rewrite_package(package, replacements, set(), args["output_path"], ".pptx")
    return {**result, "slide_count": len(slides) + len(additions), "slides_appended": len(additions)}


def reachable_parts(package: ZipFile, source: str, excluded_ids: set[str]) -> set[str]:
    """Traverse internal OPC links from the package root while skipping removed slide links.
    从包根遍历内部 OPC 链接，并跳过已删除幻灯片的链接。
    """
    # Ignore external targets and detect missing internal members before creating an output.
    # 忽略外部目标，并在生成输出前检测缺失的内部成员。
    members = set(package.namelist())
    visited = set()
    pending = [""]
    while pending:
        owner = pending.pop()
        if owner in visited:
            continue
        visited.add(owner)
        for record in relationships(package, owner):
            if record.get("TargetMode") == "External" or (owner == source and record.get("Id") in excluded_ids):
                continue
            target = relationship_target(owner, record)
            if target not in members:
                raise ValueError(f"Office relationship references a missing member: {target}.")
            if target not in visited:
                pending.append(target)
    return visited


def pptx_remove_slides(args: dict, _context: dict) -> dict:
    """Remove selected slide positions and parts no longer reachable from the presentation.
    删除选定幻灯片位置及演示文稿已无法再引用的部件。
    """
    positions = args["slides"]
    if not 1 <= len(positions) <= MAX_SLIDE_EDITS or len(set(positions)) != len(positions):
        raise ValueError(f"Provide between 1 and {MAX_SLIDE_EDITS} distinct slide positions.")
    with checked_package(read_office_zip(args["path"], ".pptx")) as package:
        # Resolve positions against the original slide list before any editing.
        # 根据原始幻灯片列表定位全部位置后再编辑。
        source, variant, presentation, document, relation_doc, types = presentation_nodes(package)
        relation = relationship_base(variant)
        lists = nodes(document.documentElement, presentation, "sldIdLst")
        if len(lists) != 1:
            raise ValueError("The presentation has no unique slide list.")
        slides = nodes(lists[0], presentation, "sldId")
        if any(not 1 <= position <= len(slides) for position in positions):
            raise ValueError("At least one selected slide position is outside the presentation.")
        selected = [slides[position - 1] for position in positions]
        identifiers = {item.getAttributeNS(relation, "id") for item in selected}
        numeric_ids = {item.getAttribute("id") for item in selected}
        parts = {target_for_id(package, source, identity, "slide") for identity in identifiers}
        # Reuse reachability to prune orphan notes and media while retaining shared resources.
        # 使用可达性清除孤立备注和媒体，同时保留共享资源。
        original = reachable_parts(package, source, set())
        remaining = reachable_parts(package, source, identifiers)
        if remaining & parts:
            raise ValueError("A selected slide remains linked from another part of the presentation.")
        discarded = original - remaining
        omitted = discarded | {rels_name(part) for part in discarded}

        # Remove references from the primary slide list, custom shows, and standard Office section lists.
        # 从主幻灯片列表、自定义放映和标准 Office 分节列表中移除引用。
        for item in selected:
            lists[0].removeChild(item)
        for item in list(document.getElementsByTagNameNS(presentation, "sld")):
            if item.getAttributeNS(relation, "id") in identifiers:
                item.parentNode.removeChild(item)
        for item in list(document.getElementsByTagNameNS(
            "http://schemas.microsoft.com/office/powerpoint/2010/main", "sldId",
        )):
            if item.getAttribute("id") in numeric_ids:
                item.parentNode.removeChild(item)
        for record in nodes(relation_doc.documentElement, REL, "Relationship"):
            if record.getAttribute("Id") in identifiers:
                record.parentNode.removeChild(record)
        for part in nodes(types.documentElement, TYPES, "Override"):
            if part.getAttribute("PartName").removeprefix("/") in discarded:
                part.parentNode.removeChild(part)
        replacements = {
            source: document.toxml(encoding="UTF-8"),
            rels_name(source): relation_doc.toxml(encoding="UTF-8"),
            "[Content_Types].xml": types.toxml(encoding="UTF-8"),
        }
        result = rewrite_package(package, replacements, omitted, args["output_path"], ".pptx")
    return {**result, "slide_count": len(slides) - len(selected), "slides_removed": len(selected)}
