param(
    [Parameter(Mandatory = $true)]
    [string]$ImagePath,
    [string]$Language = ""
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
# Resolve the one-result WinRT adapter used by image loading and OCR.
# 解析加载图像和 OCR 使用的单结果 WinRT 适配器。
$SingleResultTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name.Split([char]96)[0] -eq "IAsyncOperation"
} | Select-Object -First 1

function Await-WinRt {
    <#
    .SYNOPSIS
    Wait for an IAsyncOperation and return its typed result.
    等待 IAsyncOperation 并返回其类型化结果。
    #>
    param($Operation, [type]$ResultType)
    # Specializing AsTask bridges the WinRT operation into a synchronous script.
    # 特化 AsTask 将 WinRT 操作桥接到同步脚本。
    $Task = $SingleResultTask.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    return $Task.GetAwaiter().GetResult()
}

# Load the runtime classes with their Windows Runtime metadata.
# 通过 Windows Runtime 元数据加载运行时类型。
$Ocr = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$LanguageType = [Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
$StorageFile = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$Decoder = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$Engine = if ($Language) {
    $Ocr::TryCreateFromLanguage($LanguageType::new($Language))
} else {
    $Ocr::TryCreateFromUserProfileLanguages()
}
if ($null -eq $Engine) {
    throw "The requested Windows OCR language is not available."
}

# Read the raster and dispose of the image handles after recognition.
# 读取栅格图像，并在识别完成后释放图像句柄。
$File = Await-WinRt ($StorageFile::GetFileFromPathAsync($ImagePath)) $StorageFile
$Stream = Await-WinRt ($File.OpenAsync([Windows.Storage.FileAccessMode,Windows.Storage,ContentType=WindowsRuntime]::Read)) ([Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime])
try {
    $BitmapDecoder = Await-WinRt ($Decoder::CreateAsync($Stream)) $Decoder
    $Bitmap = Await-WinRt ($BitmapDecoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime])
    try {
        $Result = Await-WinRt ($Engine.RecognizeAsync($Bitmap)) ([Windows.Media.Ocr.OcrResult,Windows.Foundation,ContentType=WindowsRuntime])
        [Console]::Write(($Result.Lines | ForEach-Object { $_.Text }) -join [char]10)
    } finally {
        $Bitmap.Dispose()
    }
} finally {
    $Stream.Dispose()
}
