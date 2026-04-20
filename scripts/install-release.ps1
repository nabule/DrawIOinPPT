param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "Greensoft\\DrawioPpt")
)

$ErrorActionPreference = "Stop"

$sourceRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourceBin = Join-Path $sourceRoot "bin"
$targetRoot = [System.IO.Path]::GetFullPath($InstallRoot)
$targetAssemblyPath = Join-Path $targetRoot "bin\\DrawioPpt.PowerPointAddIn.dll"

if (-not (Test-Path (Join-Path $sourceBin "DrawioPpt.PowerPointAddIn.dll"))) {
    throw "Release package root not found. Expected: $sourceBin\\DrawioPpt.PowerPointAddIn.dll"
}

$runningPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
if ($runningPowerPoint) {
    throw "请先关闭正在运行的 PowerPoint，再执行安装。"
}

if (Test-Path $targetRoot) {
    Remove-Item -LiteralPath $targetRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $targetRoot | Out-Null
Copy-Item (Join-Path $sourceRoot "*") $targetRoot -Recurse -Force

$registerScript = Join-Path $targetRoot "scripts\\register-addin.ps1"
if (-not (Test-Path $registerScript)) {
    throw "Register script not found after install: $registerScript"
}

& powershell.exe -ExecutionPolicy Bypass -File $registerScript -AssemblyPath $targetAssemblyPath
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "DrawioPpt installed successfully."
Write-Host "  InstallRoot: $targetRoot"
Write-Host "  Assembly:    $targetAssemblyPath"
