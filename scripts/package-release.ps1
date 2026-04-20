param(
    [string]$Version = "v1.0.1",
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
$releaseNotesName = "release-notes-" + $Version + ".md"
$e2eDocName = "e2e-test-report-" + ($Version -replace "-stable$", "") + ".md"
$releaseEvidenceName = "release-evidence-" + $Version + ".md"
$logRoot = Join-Path $repoRoot ("artifacts\\logs\\" + $Version)
$addInBinRoot = Join-Path $repoRoot ("src\\DrawioPpt.PowerPointAddIn\\bin\\{0}\\{1}" -f $Platform, $Configuration)
$coreBinRoot = Join-Path $repoRoot ("src\\DrawioPpt.Core\\bin\\{0}\\{1}" -f $Platform, $Configuration)

if (-not $SkipBuild) {
    & powershell.exe -ExecutionPolicy Bypass -File $buildScript -Configuration $Configuration -Platform $Platform
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
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "logs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "bin\\runtimes\\win-x64\\native") | Out-Null

Copy-Item (Join-Path $addInBinRoot "DrawioPpt.PowerPointAddIn.dll") (Join-Path $packageRoot "bin") -Force
Copy-Item (Join-Path $addInBinRoot "DrawioPpt.PowerPointAddIn.pdb") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $coreBinRoot "DrawioPpt.Core.dll") (Join-Path $packageRoot "bin") -Force
Copy-Item (Join-Path $addInBinRoot "Microsoft.Web.WebView2.Core.dll") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $addInBinRoot "Microsoft.Web.WebView2.WinForms.dll") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $addInBinRoot "runtimes\\win-x64\\native\\WebView2Loader.dll") (Join-Path $packageRoot "bin\\runtimes\\win-x64\\native") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot "README.md") $packageRoot -Force
Copy-Item (Join-Path $repoRoot "docs\\installation.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\user-guide.md") (Join-Path $packageRoot "docs") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot "docs\\optimization-backlog.md") (Join-Path $packageRoot "docs") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot "docs\\regression-checklist.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $releaseNotesName)) (Join-Path $packageRoot "docs") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot ("docs\\" + $e2eDocName)) (Join-Path $packageRoot "docs") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot ("docs\\" + $releaseEvidenceName)) (Join-Path $packageRoot "docs") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot "scripts\\register-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\unregister-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\install-release.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\uninstall-release.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\url-editor-smoke.ps1") (Join-Path $packageRoot "scripts") -Force

if (Test-Path $logRoot) {
    Copy-Item (Join-Path $logRoot "*") (Join-Path $packageRoot "logs") -Recurse -Force
}

@'
@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\install-release.ps1" %*
'@ | Set-Content -Path (Join-Path $packageRoot "install.cmd") -Encoding ASCII

@'
@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall-release.ps1" %*
'@ | Set-Content -Path (Join-Path $packageRoot "uninstall.cmd") -Encoding ASCII

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
- install.cmd
- uninstall.cmd
- scripts\register-addin.ps1
- scripts\unregister-addin.ps1
- scripts\install-release.ps1
- scripts\uninstall-release.ps1
- scripts\url-editor-smoke.ps1
- docs\installation.md
- docs\user-guide.md
- docs\optimization-backlog.md
- docs\regression-checklist.md
- docs\$releaseNotesName
- docs\$e2eDocName
- docs\$releaseEvidenceName
- logs\*
"@ | Set-Content -Path $manifestPath -Encoding UTF8

Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath -Force

Write-Host "Release package created:"
Write-Host "  Folder: $packageRoot"
Write-Host "  Zip:    $zipPath"
