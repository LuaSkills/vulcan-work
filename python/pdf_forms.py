"""Fill standard AcroForm fields in a separate PDF output.
在独立的 PDF 输出中填写标准 AcroForm 字段。
"""

from contextlib import closing
from io import BytesIO
from pypdf import PdfReader, PdfWriter

from python.office import open_pdf
from python.output_files import write_new_file

# Cap the number of input fields and the length written into each field.
# 限制传入字段数量和每个字段的写入长度。
MAX_FORM_FIELDS = 100
MAX_FORM_TEXT = 10000


def pdf_fill_form(args: dict, _context: dict) -> dict:
    """Write text and choices to named PDF form fields and publish a new copy.
    向具名 PDF 表单字段写入文本或选项，并发布新副本。
    """
    # The original PDF stays closed for writing and supplies the AcroForm field tree.
    # 原始 PDF 不以写入模式打开，只提供 AcroForm 字段树。
    reader = open_pdf(args["path"])
    form = reader.root_object.get("/AcroForm")
    if form is None or "/XFA" in form.get_object():
        raise ValueError("Provide a PDF with a standard AcroForm and no XFA form.")
    # Only names declared by the document are accepted as targets.
    # 只接受文档声明的字段作为写入目标。
    declared = reader.get_fields() or {}
    if any(field.get("/FT") == "/Sig" and field.get("/V") is not None for field in declared.values()):
        raise ValueError("Signed PDFs are unsupported for form rewriting.")
    inputs = args["fields"]
    if not 1 <= len(inputs) <= MAX_FORM_FIELDS:
        raise ValueError(f"Provide between 1 and {MAX_FORM_FIELDS} fields.")
    values = {}
    for item in inputs:
        # Validate all fields before cloning the document.
        # 克隆文档前校验全部字段。
        name = item["name"]
        text = item["value"]
        if name in values or name not in declared:
            raise ValueError(f"Unknown or duplicate PDF field: {name!r}.")
        if len(text) > MAX_FORM_TEXT:
            raise ValueError(f"The value for PDF field {name!r} is too long.")
        # The field type and flags distinguish text, choice, buttons, and read-only fields.
        # 字段类型及标记区分文本、选项、按钮和只读字段。
        field = declared[name]
        kind = field.get("/FT")
        flags = int(field.get("/Ff", 0))
        if flags & 1:
            raise ValueError(f"PDF field {name!r} is read-only.")
        if kind == "/Btn" and not flags & (1 << 16):
            # PDF button states are names, including the standard /Off state.
            # PDF 按钮状态是名称，包括标准的 /Off 状态。
            text = "/" + text.removeprefix("/")
            if text not in field.get("/_States_", []):
                raise ValueError(f"Invalid button state for PDF field {name!r}: {text!r}.")
        elif kind == "/Ch":
            if text not in field.get("/_States_", []) and not flags & (1 << 18):
                raise ValueError(f"Unknown choice for PDF field {name!r}: {text!r}.")
        elif kind != "/Tx":
            raise ValueError(f"Unsupported PDF field type for {name!r}: {kind}.")
        values[name] = text

    with closing(PdfWriter(clone_from=reader)) as writer:
        writer.update_page_form_field_values(None, values, auto_regenerate=False)

        def generate(stream) -> None:
            """Serialize and verify the stored values before the output can be published.
            序列化并验证存储值后才允许发布输出。
            """
            writer.write(stream)
            # Re-read the candidate so a missing widget or unsupported update cannot silently succeed.
            # 重读候选文件，避免缺少部件或不支持的更新静默成功。
            stream.seek(0)
            updated = PdfReader(stream).get_fields() or {}
            if any(name not in updated or str(updated[name].get("/V")) != value for name, value in values.items()):
                raise ValueError("PDF form values did not round-trip in the generated document.")

        result = write_new_file(args["output_path"], ".pdf", generate)
    return {**result, "fields_written": len(values)}
