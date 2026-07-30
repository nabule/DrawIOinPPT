param(
    [string]$Version = "v1.0.10",
    [string]$Configuration = "Release",
    [string]$Platform = "x64",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildScript = Join-Path $PSScriptRoot "build.ps1"
$buildInfoScript = Join-Path $PSScriptRoot "write-build-info.ps1"
$releaseBase = [System.IO.Path]::GetFullPath(
    (Join-Path $repoRoot "artifacts\\releases"))
if ($Version -notmatch '^v[0-9A-Za-z](?:[0-9A-Za-z._-]*[0-9A-Za-z])?$') {
    throw "Version must be a single safe release name such as v1.0.8."
}

$releaseRoot = [System.IO.Path]::GetFullPath(
    (Join-Path $releaseBase $Version))
$releaseBasePrefix = $releaseBase.TrimEnd(
    [System.IO.Path]::DirectorySeparatorChar,
    [System.IO.Path]::AltDirectorySeparatorChar) +
    [System.IO.Path]::DirectorySeparatorChar
if (-not $releaseRoot.StartsWith(
        $releaseBasePrefix,
        [System.StringComparison]::OrdinalIgnoreCase) -or
    -not [string]::Equals(
        (Split-Path -Parent $releaseRoot),
        $releaseBase,
        [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Release root must be an immediate child of artifacts\\releases: $releaseRoot"
}

if (-not [string]::Equals(
        (Split-Path -Leaf $releaseRoot),
        $Version,
        [System.StringComparison]::Ordinal)) {
    throw "Version changes under Windows path normalization and is unsafe: $Version"
}

$packageRoot = Join-Path $releaseRoot "package"
$zipPath = Join-Path $releaseRoot ("DrawioPpt-" + $Version + ".zip")
$releaseNotesName = "release-notes-" + $Version + ".md"
$releaseNotesEnName = "release-notes-" + $Version + ".en.md"
$e2eDocName = "e2e-test-report-" + ($Version -replace "-stable$", "") + ".md"
$e2eDocEnName = "e2e-test-report-" + ($Version -replace "-stable$", "") + ".en.md"
$releaseEvidenceName = "release-evidence-" + $Version + ".md"
$releaseEvidenceEnName = "release-evidence-" + $Version + ".en.md"
$stressReportName = "word-ui-thread-stress-report-" + $Version + ".md"
$stressReportEnName = "word-ui-thread-stress-report-" + $Version + ".en.md"
$stressEvidenceName = "word-ui-thread-stress-" + $Version + ".json"
$addInBinRoot = Join-Path $repoRoot ("src\\DrawioPpt.PowerPointAddIn\\bin\\{0}\\{1}" -f $Platform, $Configuration)
$wordAddInBinRoot = Join-Path $repoRoot ("src\\DrawioPpt.WordAddIn\\bin\\{0}\\{1}" -f $Platform, $Configuration)
$coreBinRoot = Join-Path $repoRoot ("src\\DrawioPpt.Core\\bin\\{0}\\{1}" -f $Platform, $Configuration)

function Assert-PackageMarkdownLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot
    )

    $normalizedPackageRoot = [System.IO.Path]::GetFullPath($PackageRoot)
    $packagePrefix = $normalizedPackageRoot.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $missingLinks = New-Object System.Collections.Generic.List[string]
    $markdownFiles = @(Get-ChildItem -LiteralPath $normalizedPackageRoot -Recurse -Filter "*.md" -File)
    foreach ($markdownFile in $markdownFiles) {
        $sourceText = Get-Content -Raw -LiteralPath $markdownFile.FullName
        $matches = [regex]::Matches($sourceText, '\[[^\]]*\]\(([^)]+)\)')
        foreach ($match in $matches) {
            $target = $match.Groups[1].Value.Trim()
            if ([string]::IsNullOrWhiteSpace($target) -or
                $target.StartsWith("#", [System.StringComparison]::Ordinal) -or
                $target -match '^[a-zA-Z][a-zA-Z0-9+.-]*:') {
                continue
            }

            $pathPart = ($target -split '[#?]', 2)[0].Trim().Trim("<", ">")
            if ([string]::IsNullOrWhiteSpace($pathPart)) {
                continue
            }

            $decodedPath = [Uri]::UnescapeDataString($pathPart).Replace(
                [System.IO.Path]::AltDirectorySeparatorChar,
                [System.IO.Path]::DirectorySeparatorChar)
            $resolvedTarget = [System.IO.Path]::GetFullPath(
                (Join-Path $markdownFile.DirectoryName $decodedPath))
            $isInsidePackage = $resolvedTarget.StartsWith(
                $packagePrefix,
                [System.StringComparison]::OrdinalIgnoreCase)
            if (-not $isInsidePackage -or -not (Test-Path -LiteralPath $resolvedTarget)) {
                $relativeSource = $markdownFile.FullName.Substring($packagePrefix.Length)
                $missingLinks.Add("$relativeSource -> $target") | Out-Null
            }
        }
    }

    foreach ($missingLink in $missingLinks) {
        Write-Host "PackageMarkdownMissingLink=$missingLink"
    }

    Write-Host "PackageMarkdownMissingLinkCount=$($missingLinks.Count)"
    if ($missingLinks.Count -gt 0) {
        throw "Release package contains $($missingLinks.Count) unresolved local Markdown link(s)."
    }
}

function New-PortableZip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $sourceRootFullPath = [System.IO.Path]::GetFullPath($SourceRoot)
    $sourcePrefix = $sourceRootFullPath.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $destinationFullPath = [System.IO.Path]::GetFullPath($DestinationPath)

    if (Test-Path -LiteralPath $destinationFullPath) {
        Remove-Item -LiteralPath $destinationFullPath -Force
    }

    $archiveStream = [System.IO.File]::Open(
        $destinationFullPath,
        [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None)
    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipArchive]::new(
            $archiveStream,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false)
        foreach ($file in Get-ChildItem -LiteralPath $sourceRootFullPath -Recurse -File | Sort-Object FullName) {
            $entryName = $file.FullName.Substring($sourcePrefix.Length).Replace(
                [System.IO.Path]::DirectorySeparatorChar,
                [System.IO.Path]::AltDirectorySeparatorChar)
            $entry = $archive.CreateEntry(
                $entryName,
                [System.IO.Compression.CompressionLevel]::Optimal)
            $entryStream = $null
            $sourceStream = $null
            try {
                $entryStream = $entry.Open()
                $sourceStream = [System.IO.File]::OpenRead($file.FullName)
                $sourceStream.CopyTo($entryStream)
            }
            finally {
                if ($sourceStream -ne $null) {
                    $sourceStream.Dispose()
                }
                if ($entryStream -ne $null) {
                    $entryStream.Dispose()
                }
            }
        }
    }
    finally {
        if ($archive -ne $null) {
            $archive.Dispose()
        }
        $archiveStream.Dispose()
    }
}

function Assert-PortableZip {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$ZipPath,

        [Parameter(Mandatory = $true)]
        [string]$VerificationRoot
    )

    $sourceRootFullPath = [System.IO.Path]::GetFullPath($SourceRoot)
    $zipFullPath = [System.IO.Path]::GetFullPath($ZipPath)
    $verificationFullPath = [System.IO.Path]::GetFullPath($VerificationRoot)
    $archive = [System.IO.Compression.ZipFile]::OpenRead($zipFullPath)
    try {
        $fileEntries = @($archive.Entries | Where-Object { -not [string]::IsNullOrEmpty($_.Name) })
        $invalidEntries = @($fileEntries | Where-Object {
                $_.FullName.Contains("\") -or
                $_.FullName.StartsWith("/") -or
                $_.FullName -match '(^|/)\.\.(/|$)'
            })
        if ($invalidEntries.Count -gt 0) {
            throw "Portable ZIP contains invalid entry names: $($invalidEntries.FullName -join ', ')"
        }

        $sourceFiles = @(Get-ChildItem -LiteralPath $sourceRootFullPath -Recurse -File)
        if ($fileEntries.Count -ne $sourceFiles.Count) {
            throw "Portable ZIP entry count $($fileEntries.Count) does not match package file count $($sourceFiles.Count)."
        }
    }
    finally {
        $archive.Dispose()
    }

    if (Test-Path -LiteralPath $verificationFullPath) {
        Remove-Item -LiteralPath $verificationFullPath -Recurse -Force
    }

    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory(
            $zipFullPath,
            $verificationFullPath)
        $sourcePrefix = $sourceRootFullPath.TrimEnd(
            [System.IO.Path]::DirectorySeparatorChar,
            [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
        foreach ($sourceFile in Get-ChildItem -LiteralPath $sourceRootFullPath -Recurse -File) {
            $relativePath = $sourceFile.FullName.Substring($sourcePrefix.Length)
            $extractedPath = Join-Path $verificationFullPath $relativePath
            if (-not (Test-Path -LiteralPath $extractedPath)) {
                throw "Portable ZIP extraction is missing: $relativePath"
            }

            $sourceHash = (Get-FileHash -LiteralPath $sourceFile.FullName -Algorithm SHA256).Hash
            $extractedHash = (Get-FileHash -LiteralPath $extractedPath -Algorithm SHA256).Hash
            if (-not [string]::Equals(
                    $sourceHash,
                    $extractedHash,
                    [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Portable ZIP extracted hash mismatch: $relativePath"
            }
        }
    }
    finally {
        if (Test-Path -LiteralPath $verificationFullPath) {
            Remove-Item -LiteralPath $verificationFullPath -Recurse -Force
        }
    }

    Write-Host "PortableZipEntryCount=$($sourceFiles.Count)"
    Write-Host "PortableZipForwardSlashEntries=True"
    Write-Host "PortableZipExtractedHashesMatch=True"
}

if (-not $SkipBuild) {
    & powershell.exe -ExecutionPolicy Bypass -File $buildScript -Configuration $Configuration -Platform $Platform -Version $Version
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if (-not (Test-Path $addInBinRoot)) {
    throw "Add-in output folder not found: $addInBinRoot"
}

if (-not (Test-Path $wordAddInBinRoot)) {
    throw "Word add-in output folder not found: $wordAddInBinRoot"
}

if (-not (Test-Path $coreBinRoot)) {
    throw "Core output folder not found: $coreBinRoot"
}

if (-not (Test-Path $buildInfoScript)) {
    throw "Build info script not found: $buildInfoScript"
}

$buildInfoOutputRoots = @($addInBinRoot, $wordAddInBinRoot, $coreBinRoot)
& $buildInfoScript -Version $Version -OutputRoots $buildInfoOutputRoots

$buildInfoPath = Join-Path $addInBinRoot "BuildInfo.txt"
if (-not (Test-Path -LiteralPath $buildInfoPath)) {
    throw "BuildInfo.txt was not generated: $buildInfoPath"
}

$buildInfoValues = @{}
foreach ($line in Get-Content -LiteralPath $buildInfoPath) {
    if ($line -match '^\s*([^:=]+)\s*[:=]\s*(.*)\s*$') {
        $buildInfoValues[$Matches[1].Trim()] = $Matches[2].Trim()
    }
}

$buildId = $buildInfoValues["BuildId"]
$gitShortHash = $buildInfoValues["GitShortHash"]
if ([string]::IsNullOrWhiteSpace($buildId)) {
    $buildId = $Version
}

if ([string]::IsNullOrWhiteSpace($gitShortHash)) {
    $gitShortHash = "unknown"
}

if (Test-Path $releaseRoot) {
    Remove-Item -LiteralPath $releaseRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $packageRoot | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "docs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "docs\\evidence") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "scripts") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "bin") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $packageRoot "bin\\runtimes\\win-x64\\native") | Out-Null

Copy-Item (Join-Path $addInBinRoot "DrawioPpt.PowerPointAddIn.dll") (Join-Path $packageRoot "bin") -Force
Copy-Item (Join-Path $wordAddInBinRoot "DrawioPpt.WordAddIn.dll") (Join-Path $packageRoot "bin") -Force
Copy-Item (Join-Path $coreBinRoot "DrawioPpt.Core.dll") (Join-Path $packageRoot "bin") -Force
Copy-Item $buildInfoPath $packageRoot -Force
Copy-Item $buildInfoPath (Join-Path $packageRoot "bin") -Force
Copy-Item (Join-Path $addInBinRoot "Microsoft.Web.WebView2.Core.dll") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $addInBinRoot "Microsoft.Web.WebView2.WinForms.dll") (Join-Path $packageRoot "bin") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $addInBinRoot "runtimes\\win-x64\\native\\WebView2Loader.dll") (Join-Path $packageRoot "bin\\runtimes\\win-x64\\native") -Force -ErrorAction SilentlyContinue
Copy-Item (Join-Path $repoRoot "README.md") $packageRoot -Force
Copy-Item (Join-Path $repoRoot "README.en.md") $packageRoot -Force
Copy-Item (Join-Path $repoRoot "docs\\installation.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\installation.en.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\architecture.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\architecture.en.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\user-guide.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\user-guide.en.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\word-addin.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\word-addin.en.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\optimization-backlog.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\optimization-backlog.en.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\regression-checklist.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\regression-checklist.en.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\e2e-test-report-v1.0.5.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot "docs\\e2e-test-report-v1.0.5.en.md") (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $releaseNotesName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $releaseNotesEnName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $e2eDocName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $e2eDocEnName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $releaseEvidenceName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $releaseEvidenceEnName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $stressReportName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\" + $stressReportEnName)) (Join-Path $packageRoot "docs") -Force
Copy-Item (Join-Path $repoRoot ("docs\\evidence\\" + $stressEvidenceName)) (Join-Path $packageRoot "docs\\evidence") -Force
Copy-Item (Join-Path $repoRoot "scripts\\register-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\unregister-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\register-word-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\unregister-word-addin.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\register-office-addins.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\unregister-office-addins.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\verify-office-install.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-addin-load-check.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-selection-event-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-url-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-url-addin-host-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-complex-metadata-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-ui-thread-stress-acceptance.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-ui-thread-rendering-validation-test.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-svg-aspect-ratio-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-preview-image-provider-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\powerpoint-svg-aspect-ratio-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\webview2-user-data-folder-test.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\word-url-addin-host-e2e-safety-test.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\version-display-safety-test.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\install-release-upgrade-safety-test.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\install-release-encoding-safety-test.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\write-build-info.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\install-release.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\uninstall-release.ps1") (Join-Path $packageRoot "scripts") -Force
Copy-Item (Join-Path $repoRoot "scripts\\url-editor-smoke.ps1") (Join-Path $packageRoot "scripts") -Force

@'
@echo off
setlocal
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\install-release.ps1" %*
set "exitCode=%ERRORLEVEL%"
if not "%exitCode%"=="0" (
    echo.
    echo DrawioPpt installation failed with exit code %exitCode%.
    echo Review the error messages above, close Word and PowerPoint, then try again.
)
exit /b %exitCode%
'@ | Set-Content -Path (Join-Path $packageRoot "install.cmd") -Encoding ASCII

@'
@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall-release.ps1" %*
'@ | Set-Content -Path (Join-Path $packageRoot "uninstall.cmd") -Encoding ASCII

$manifestPath = Join-Path $packageRoot "PACKAGE.txt"
@"
Package: DrawioPpt
Version: $Version
BuildId: $buildId
GitShortHash: $gitShortHash
Configuration: $Configuration
Platform: $Platform
BuiltAt: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

Contents:
- README.md
- README.en.md
- PACKAGE.txt
- BuildInfo.txt
- bin\DrawioPpt.PowerPointAddIn.dll
- bin\DrawioPpt.WordAddIn.dll
- bin\DrawioPpt.Core.dll
- bin\BuildInfo.txt
- bin\Microsoft.Web.WebView2.Core.dll (if available)
- bin\Microsoft.Web.WebView2.WinForms.dll (if available)
- bin\runtimes\win-x64\native\WebView2Loader.dll (if available)
- install.cmd
- uninstall.cmd
- scripts\register-addin.ps1
- scripts\unregister-addin.ps1
- scripts\register-word-addin.ps1
- scripts\unregister-word-addin.ps1
- scripts\register-office-addins.ps1
- scripts\unregister-office-addins.ps1
- scripts\verify-office-install.ps1
- scripts\word-addin-load-check.ps1
- scripts\word-selection-event-e2e.ps1
- scripts\word-url-e2e.ps1
- scripts\word-url-addin-host-e2e.ps1
- scripts\word-complex-metadata-e2e.ps1
- scripts\word-ui-thread-stress-acceptance.ps1
- scripts\word-ui-thread-rendering-validation-test.ps1
- scripts\word-svg-aspect-ratio-e2e.ps1
- scripts\word-preview-image-provider-e2e.ps1
- scripts\powerpoint-svg-aspect-ratio-e2e.ps1
- scripts\webview2-user-data-folder-test.ps1
- scripts\word-url-addin-host-e2e-safety-test.ps1
- scripts\version-display-safety-test.ps1
- scripts\install-release-upgrade-safety-test.ps1
- scripts\install-release-encoding-safety-test.ps1
- scripts\write-build-info.ps1
- scripts\install-release.ps1
- scripts\uninstall-release.ps1
- scripts\url-editor-smoke.ps1
- docs\installation.md
- docs\installation.en.md
- docs\architecture.md
- docs\architecture.en.md
- docs\user-guide.md
- docs\user-guide.en.md
- docs\word-addin.md
- docs\word-addin.en.md
- docs\optimization-backlog.md
- docs\optimization-backlog.en.md
- docs\regression-checklist.md
- docs\regression-checklist.en.md
- docs\e2e-test-report-v1.0.5.md
- docs\e2e-test-report-v1.0.5.en.md
- docs\$releaseNotesName
- docs\$releaseNotesEnName
- docs\$e2eDocName
- docs\$e2eDocEnName
- docs\$releaseEvidenceName
- docs\$releaseEvidenceEnName
- docs\$stressReportName
- docs\$stressReportEnName
- docs\evidence\$stressEvidenceName
"@ | Set-Content -Path $manifestPath -Encoding UTF8

Assert-PackageMarkdownLinks -PackageRoot $packageRoot

New-PortableZip -SourceRoot $packageRoot -DestinationPath $zipPath
Assert-PortableZip `
    -SourceRoot $packageRoot `
    -ZipPath $zipPath `
    -VerificationRoot (Join-Path $releaseRoot ".zip-verification")

Write-Host "Release package created:"
Write-Host "  Folder: $packageRoot"
Write-Host "  Zip:    $zipPath"
