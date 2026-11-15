<#
.SYNOPSIS
Run the office call tests, validate the skill, and build its standard package.
执行办公工具调用测试、校验技能并生成标准发布包。
#>
$ErrorActionPreference = "Stop"
# Anchor the working directory to the skill package.
# 将工作目录固定在技能包根目录。
$ProjectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
# Reuse the official wrapper for runtime loading, validation, and packaging.
# 复用官方包装器加载运行时、校验技能和打包。
$DebugScript = Join-Path $PSScriptRoot "debug_skill.ps1"
Push-Location $ProjectRoot
try {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $DebugScript -Command inspect -Output json
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & powershell -NoProfile -ExecutionPolicy Bypass -File $DebugScript -Command list-tools -Output content
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & python -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & powershell -NoProfile -ExecutionPolicy Bypass -File $DebugScript -Command validate
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    & powershell -NoProfile -ExecutionPolicy Bypass -File $DebugScript -Command package
    exit $LASTEXITCODE
} finally {
    Pop-Location
}
