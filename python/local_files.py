"""Read bounded local files and Office ZIP containers.
读取大小受限的本地文件和 Office ZIP 容器。
"""

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

# Limit bytes read from the local file before a parser sees it.
# 在解析前限制从本地文件读取的字节数。
MAX_INPUT_BYTES = 32 * 1024 * 1024
# Limit the Office ZIP container's advertised expanded size before reading worksheets.
# 在读取工作表前限制 Office ZIP 容器声明的解压大小。
MAX_OFFICE_EXPANDED_BYTES = 128 * 1024 * 1024

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


def read_office_zip(path: str, extension: str) -> bytes:
    """Return the ZIP snapshot at path after verifying its expanded-size budget.
    校验 path 对应 ZIP 快照的解压体积后，返回该快照。
    """
    # Share the exact same snapshot between the archive guard and the consumer.
    # 归档检查与后续读取使用同一份快照。
    content = read_local_file(path, extension)
    with ZipFile(BytesIO(content)) as archive:
        validate_expanded_size(archive, extension)
    return content


def validate_expanded_size(archive: ZipFile, extension: str) -> None:
    """Apply the same expanded-size budget to a source or generated Office ZIP.
    对输入与生成的 Office ZIP 使用相同的解压大小预算。
    """
    if sum(member.file_size for member in archive.infolist()) > MAX_OFFICE_EXPANDED_BYTES:
        raise ValueError(f"Expanded {extension[1:].upper()} container exceeds 128 MiB.")
