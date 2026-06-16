param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "Greensoft\\DrawioPpt")
)

$ErrorActionPreference = "Stop"

$targetRoot = [System.IO.Path]::GetFullPath($InstallRoot)
$runningPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
if ($runningPowerPoint) {
    throw "请先关闭正在运行的 PowerPoint，再执行卸载。"
}

$runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
if ($runningWord) {
    throw "请先关闭正在运行的 Word，再执行卸载。"
}

$unregisterScript = Join-Path $targetRoot "scripts\\unregister-office-addins.ps1"
if (Test-Path $unregisterScript) {
    & powershell.exe -ExecutionPolicy Bypass -File $unregisterScript
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if (Test-Path $targetRoot) {
    Set-Location $env:TEMP
    Remove-Item -LiteralPath $targetRoot -Recurse -Force
}

Write-Host "DrawioPpt uninstalled successfully."
Write-Host "  InstallRoot: $targetRoot"
