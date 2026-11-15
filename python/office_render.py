"""Render Office previews and recognize scanned PDF pages using local tools.
使用本地工具渲染 Office 预览，并识别扫描 PDF 页。
"""

from contextlib import contextmanager
from io import BytesIO
from os import environ
from pathlib import Path
from shutil import rmtree
from subprocess import run
from sys import platform
from tempfile import TemporaryDirectory
from time import monotonic, sleep

from pypdf import PdfReader

from python.local_files import read_local_file, read_office_zip
from python.output_files import write_new_file
from python.text_pages import document_text, single_item_offset

# Cap each OCR call and keep its child processes inside the Lua worker deadline.
# 限制每次 OCR 的页数，并将子进程控制在 Lua 工作器的截止时间内。
MAX_OCR_PAGES = 3
RENDER_TIMEOUT = 45
WINDOWS_LANGUAGES = {"eng": "en-US", "chi_sim": "zh-Hans-CN", "chi_tra": "zh-Hant-TW", "auto": ""}


def execute_program(command: list[str], timeout: float, cwd: Path | None = None,
                    env: dict[str, str] | None = None) -> str:
    """Run a known local office executable and return its UTF-8 output.
    运行已知本地办公程序并返回 UTF-8 输出。
    """
    # Fixed argument arrays avoid shell interpretation of file paths.
    # 固定参数数组避免文件路径被 shell 解释。
    result = run(command, capture_output=True, encoding="utf-8", errors="replace",
                 timeout=timeout, cwd=cwd, env=env)
    if result.returncode:
        diagnostic = (result.stderr + result.stdout).strip()[-1000:]
        raise ValueError(f"{Path(command[0]).name} exited with {result.returncode}: {diagnostic}")
    return result.stdout


def libreoffice_binary(resolved: str | None) -> str:
    """Find LibreOffice on PATH or in its normal desktop installation directory.
    在 PATH 或标准桌面安装目录中定位 LibreOffice。
    """
    # Windows and macOS installers commonly expose the executable only inside the installation directory.
    # Windows 与 macOS 的安装程序通常只在安装目录中暴露可执行文件。
    if resolved:
        return resolved
    candidate = None
    if platform == "win32":
        # Windows registers the system Program Files directory even in a sanitized worker.
        # 即使工作器环境被净化，Windows 仍在注册表中记录 Program Files 目录。
        from winreg import HKEY_LOCAL_MACHINE, KEY_READ, KEY_WOW64_64KEY, OpenKey, QueryValueEx
        with OpenKey(HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion",
                     access=KEY_READ | KEY_WOW64_64KEY) as key:
            directory, _ = QueryValueEx(key, "ProgramFilesDir")
        candidate = Path(directory) / "LibreOffice" / "program" / "soffice.com"
    elif platform == "darwin":
        candidate = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if candidate is not None and candidate.is_file():
        return str(candidate)
    raise ValueError("LibreOffice is required for this operation. Install it or add soffice to PATH.")


@contextmanager
def render_workspace(directory: str | None):
    """Clean up a private render workspace after LibreOffice has released its files.
    在 LibreOffice 释放文件后清理私有渲染工作区。
    """
    # A compact path leaves room for LibreOffice's nested registry and cache directories.
    # 使用简短路径，为 LibreOffice 的嵌套注册表和缓存目录预留空间。
    temporary = TemporaryDirectory(prefix="vw-", dir=directory, ignore_cleanup_errors=True)
    try:
        yield Path(temporary.name)
    finally:
        # On Windows the helper can finish updating its profile just after the launcher exits.
        # Windows 上辅助进程可能在启动器退出后才完成配置更新。
        for _ in range(5):
            try:
                rmtree(temporary.name)
                break
            except FileNotFoundError:
                break
            except OSError:
                sleep(0.1)
        temporary.cleanup()


def convert_with_libreoffice(source: Path, folder: Path, target_format: str, resolved: str | None) -> Path:
    """Convert a staged file with an isolated LibreOffice user profile.
    使用隔离的 LibreOffice 用户配置转换暂存文件。
    """
    # Isolate the profile so only this invocation handles the document.
    # 隔离用户配置，确保文档只由本次调用处理。
    profile = (folder / "p").as_uri()
    conversion = folder / "out"
    conversion.mkdir()
    # Restrict the application's working directory and temporary files to the private workspace.
    # 将应用程序的工作目录和临时文件限制在私有工作区。
    environment = dict(environ)
    for name in ("TMPDIR", "TEMP", "TMP"):
        environment[name] = str(folder)
    execute_program([libreoffice_binary(resolved), f"-env:UserInstallation={profile}", "--headless",
                     "--convert-to", target_format, "--outdir", str(conversion), str(source)],
                    RENDER_TIMEOUT, cwd=folder, env=environment)
    # LibreOffice can return zero when a filter could not create its output.
    # 转换过滤器未生成文件时，LibreOffice 也可能返回零。
    output = conversion / source.with_suffix("." + target_format.split(":", 1)[0]).name
    if not output.is_file():
        raise ValueError("LibreOffice did not create the requested output.")
    return output


def pptx_export_preview(args: dict, _context: dict) -> dict:
    """Render a presentation as a PDF preview with LibreOffice Impress.
    使用 LibreOffice Impress 将演示文稿渲染为 PDF 预览。
    """
    with render_workspace(args.get("_render_temp")) as folder:
        # Stage the bounded input so LibreOffice never opens the user's original file.
        # 暂存大小受限的输入，避免 LibreOffice 打开用户原件。
        source = folder / "slides.pptx"
        source.write_bytes(read_office_zip(args["path"], ".pptx"))
        preview = convert_with_libreoffice(source, folder, "pdf:impress_pdf_Export", args.get("_soffice"))
        content = read_local_file(str(preview), ".pdf")
    return write_new_file(args["output_path"], ".pdf", lambda stream: stream.write(content))


def recognize_image(image: Path, language: str, timeout: float, tesseract: str | None, powershell: str | None) -> str:
    """Run Tesseract where installed, or the Windows OCR engine on Windows.
    在已安装时运行 Tesseract，Windows 上也可使用系统 OCR 引擎。
    """
    # Language names come from the public schema and map to installed engine languages.
    # 语言名来自公开 Schema，并映射到各引擎安装的语言。
    if language not in WINDOWS_LANGUAGES:
        raise ValueError("Unsupported OCR language.")
    if tesseract:
        return execute_program([tesseract, str(image), "stdout", "-l", "eng" if language == "auto" else language], timeout)
    if platform == "win32":
        # Use the packaged WinRT adapter and Windows language resources.
        # 使用随包提供的 WinRT 适配器和 Windows 语言资源。
        if not powershell:
            raise ValueError("Windows PowerShell is required for Windows OCR.")
        # Pass a private copy because the managed package snapshot is held open by the engine.
        # 受管包快照由引擎保持打开，因此向 PowerShell 传递私有副本。
        script = image.parent / "ocr.ps1"
        script.write_bytes(Path(__file__).with_name("ocr_windows.ps1").read_bytes())
        command = [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
                   "-ImagePath", str(image)]
        if language != "auto":
            command.extend(["-Language", WINDOWS_LANGUAGES[language]])
        return execute_program(command, timeout)
    raise ValueError("Tesseract is required for OCR on this platform.")


def pdf_ocr_pages(args: dict, _context: dict) -> dict:
    """Recognize a small PDF page range and return resumable page text.
    识别小范围 PDF 页面并返回支持续读的页面文本。
    """
    # Validate the page selection against the actual PDF before invoking rendering.
    # 调用渲染前，根据实际 PDF 校验页码选区。
    start, end = args["start_page"], args["end_page"]
    content = read_local_file(args["path"], ".pdf")
    reader = PdfReader(BytesIO(content))
    if reader.is_encrypted:
        raise ValueError("Encrypted PDF files are unsupported.")
    page_count = len(reader.pages)
    if not 1 <= start <= end <= page_count or end - start + 1 > MAX_OCR_PAGES:
        raise ValueError(f"Select up to {MAX_OCR_PAGES} pages within this PDF's {page_count} pages.")
    offset = single_item_offset(args.get("text_offset", 0), end - start + 1)
    rasterizer = args.get("_rasterizer")
    if not rasterizer:
        raise ValueError("Poppler pdftoppm is required for PDF OCR.")
    # A shared deadline bounds multiple rendering and OCR processes in one call.
    # 多个渲染和识别子进程共享单次调用的截止时间。
    deadline = monotonic() + RENDER_TIMEOUT
    pages = []
    with TemporaryDirectory(prefix="vulcan-work-ocr-") as directory:
        # Render a bounded local snapshot, leaving the user's original path untouched.
        # 渲染大小受限的本地快照，避免接触用户源路径。
        folder = Path(directory)
        source = folder / "source.pdf"
        source.write_bytes(content)
        for page in range(start, end + 1):
            # Bound pixel dimensions for memory and the Windows OCR API.
            # 限制像素尺寸，满足内存预算与 Windows OCR API 限制。
            image_base = folder / f"page-{page}"
            execute_program([rasterizer, "-f", str(page), "-l", str(page), "-singlefile",
                             "-scale-to", "2400", "-png", str(source), str(image_base)],
                            max(0.01, deadline - monotonic()))
            text = recognize_image(image_base.with_suffix(".png"), args.get("language", "auto"),
                                   max(0.01, deadline - monotonic()),
                                   args.get("_tesseract"), args.get("_powershell"))
            pages.append({"page": page, **document_text(text.strip(), offset)})
    return {"page_count": page_count, "pages": pages}
