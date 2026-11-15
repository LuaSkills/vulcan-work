"""Publish newly generated office files without replacing an existing path.
发布新生成的办公文件，并保护已有路径。
"""

from collections.abc import Callable
from os import fsync, link, name, path as os_path, rename
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import BinaryIO
from zipfile import ZipFile

from python.local_files import MAX_INPUT_BYTES, validate_expanded_size


def write_new_file(raw_path: str, extension: str, generate: Callable[[BinaryIO], object]) -> dict:
    """Generate a file in a sibling temporary path, then publish it exclusively.
    在同目录临时路径生成文件，再以独占方式发布。
    """
    # This skill writes only full local paths in an existing output directory.
    # 本技能仅写入已有目录中的本地绝对路径。
    destination = Path(raw_path)
    if not destination.is_absolute() or destination.suffix.lower() != extension or not destination.parent.is_dir():
        raise ValueError(f"Provide an absolute output path ending in {extension} in an existing directory.")
    if os_path.lexists(destination):
        raise FileExistsError(f"Output file already exists: {destination}")
    # Keep the temporary file on the destination volume so publication is atomic.
    # 将临时文件放在目标卷上，确保发布操作具有原子性。
    temporary = NamedTemporaryFile(mode="w+b", prefix=".vulcan-work-", suffix=".tmp",
                                   dir=destination.parent, delete=False)
    try:
        with temporary:
            generate(temporary)
            # Enforce the same size budget used by the read tools.
            # 使用与读取工具相同的文件大小上限。
            size = temporary.seek(0, 2)
            if not 0 < size <= MAX_INPUT_BYTES:
                raise ValueError("Generated file is empty or exceeds the 32 MiB output limit.")
            temporary.flush()
            # The read tools also check expanded Office size, so generated ZIPs must pass it before publication.
            # 读取工具也会检查 Office 解压大小，因此发布前生成的 ZIP 必须通过相同校验。
            if extension != ".pdf":
                with ZipFile(temporary) as archive:
                    validate_expanded_size(archive, extension)
            fsync(temporary.fileno())
        # Windows rename fails if the destination exists; POSIX link provides the same exclusive publication.
        # Windows 重命名会拒绝已有目标，POSIX 硬链接提供同样的独占发布语义。
        try:
            if name == "nt":
                rename(temporary.name, destination)
            else:
                link(temporary.name, destination)
        except FileExistsError as error:
            raise FileExistsError(f"Output file already exists: {destination}") from error
    finally:
        Path(temporary.name).unlink(missing_ok=True)
    return {"output_path": str(destination), "size_bytes": size}
