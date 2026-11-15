"""Paginate bounded text and ordered items without losing content.
在保留完整内容的前提下，为文本与有序条目分页。
"""

# A shared limit keeps document text chunks within the same output budget.
# 统一上限让各种文档的文本段使用同一返回预算。
DOCUMENT_TEXT_CHARS = 1000


def select_items(items: list, start: int, limit: int, maximum: int) -> tuple[list, int | None]:
    """Return a bounded one-based selection and the next index, if present.
    返回一基范围内的限定条目，有后续时给出下一项索引。
    """
    if not 1 <= limit <= maximum:
        raise ValueError(f"Select between 1 and {maximum} items per call.")
    if not 1 <= start <= max(1, len(items)):
        raise ValueError(f"Start index {start} is outside the {len(items)} available items.")
    # The exclusive slice end also counts items consumed from the start of the list.
    # 切片结束索引也是从列表起点已读取的条目数。
    end = min(start + limit - 1, len(items))
    return items[start - 1:end], end + 1 if end < len(items) else None


def single_item_offset(offset: int, item_count: int) -> int:
    """Validate a character offset for an exact single-item selection.
    为单条内容的确切选区校验字符偏移。
    """
    if offset < 0 or (offset > 0 and item_count != 1):
        raise ValueError("Use a nonnegative character offset and select exactly one item to resume text.")
    return offset


def slice_text(text: str, offset: int, limit: int) -> tuple[str, int | None]:
    """Return text starting at a Unicode character offset and its next offset.
    从 Unicode 字符偏移处返回文本，并给出下一段偏移。
    """
    # Use character indices so multilingual text never splits a UTF-8 sequence.
    # 使用字符索引，避免截断多语言文本中的 UTF-8 字符序列。
    end = offset + limit
    return text[offset:end], end if end < len(text) else None


def document_text(text: str, offset: int) -> dict:
    """Return a document text chunk and its continuation information.
    返回文档文本段及其续读信息。
    """
    # The continuation starts immediately after the returned characters.
    # 续读从返回字符之后的位置开始。
    content, next_offset = slice_text(text, offset, DOCUMENT_TEXT_CHARS)
    return {"text": content, "text_truncated": next_offset is not None, "next_text_offset": next_offset}
