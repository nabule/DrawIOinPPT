param(
    [string]$Version = "v0.6.0-stable",
    [string]$Configuration = "Release",
    [string]$Platform = "x64",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$releaseRoot = Join-Path $repoRoot ("artifacts\\releases\\" + $Version)
$packageRoot = Join-Path $releaseRoot "package"
$zipPath = Join-Path $releaseRoot ("DrawioPpt-" + $Version + ".zip")
$addInBinRoot = Join-Path $repoRoot ("src\\DrawioPpt.PowerPointAddIn\\bin\\{0}\\{1}" -f $Platform, $Configuration)
$coreBinRoot = Join-Path $repoRoot ("src\\DrawioPpt.Core\\bin\\{0}\\{1}" -f $Platform, $Configuration)

if (-not $SkipBuild) {
    & powershell -ExecutionPolicy Bypass -File $buildScript -Configuration $Configuration -Platform $Platform
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if (-not (Test-Path $addInBinRoot)) {
    throw "Add-in output folder not found: $addInBinRoot"
}

if (-not (Test-Path $coreBinRoot)) {
    throw "Core output folder not found: $coreBinRoot"
}

if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $packageRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "docs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "scripts") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "bin") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "bin\\runtimes\\win-x64\\native") | Out-Null

Copy-Item (Join-Path $addInBinRoot "DrawioPpt.PowerPointAddIn.dll") (Join-Path $packageRoot "bin") -Force
Copy-Item (Join-Path $addInBinRoot "DrawioPpt.PowerPointAddIn.pdb") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $coreBinRoot "DrawioPpt.Core.dll") (Join-Path $packageRoot "bin") -Force
Copy-Item (Join-Path $addInBinRoot "Microsoft.Web.WebView2.Core.dll") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $addInBinRoot "Microsoft.Web.WebView2.WinForms.dll") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $addInBinRoot "runtimes\\win-x64\\native\\WebView2Loader.dll") (Join-Path $packageRoot "bin\\runtimes\\win-x64\\native") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot "README.md") $packageRoot -Force
Copy-Item (Join-Path $repoRoot "docs\\installation.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\regression-checklist.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\release-notes-v0.6.0-stable.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "scripts\\register-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\unregister-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\url-editor-smoke.ps1") (Join-Path $packageRoot "scripts") -Force

$manifestPath = Join-Path $packageRoot "PACKAGE.txt"
@"
Package: DrawioPpt
Version: $Version
Configuration: $Configuration
Platform: $Platform
BuiltAt: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

Contents:
- bin\DrawioPpt.PowerPointAddIn.dll
- bin\DrawioPpt.Core.dll
- bin\Microsoft.Web.WebView2.Core.dll (if available)
- bin\Microsoft.Web.WebView2.WinForms.dll (if available)
- bin\runtimes\win-x64\native\WebView2Loader.dll (if available)
- scripts\register-addin.ps1
- scripts\unregister-addin.ps1
- scripts\url-editor-smoke.ps1
- docs\installation.md
- docs\regression-checklist.md
- docs\release-notes-v0.6.0-stable.md
"@ | Set-Content -Path $manifestPath -Encoding UTF8

Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force

Write-Host "Release package created:"
Write-Host "  Folder: $packageRoot"
Write-Host "  Zip:    $zipPath"
