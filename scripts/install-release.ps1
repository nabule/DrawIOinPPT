param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "Greensoft\\DrawioPpt")
)

$ErrorActionPreference = "Stop"

$sourceRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourceBin = Join-Path $sourceRoot "bin"
$targetRoot = [System.IO.Path]::GetFullPath($InstallRoot)
$targetAssemblyPath = Join-Path $targetRoot "bin\\DrawioPpt.PowerPointAddIn.dll"
$targetWordAssemblyPath = Join-Path $targetRoot "bin\\DrawioPpt.WordAddIn.dll"
$stagingRoot = Join-Path $env:TEMP ("DrawioPpt\\install-stage-" + [Guid]::NewGuid().ToString("N"))

function Remove-ExistingInstallRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return
    }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force
            return
        }
        catch {
            if ($attempt -lt 3) {
                Start-Sleep -Seconds 2
                continue
            }

            $webView2 = Get-Process -Name msedgewebview2 -ErrorAction SilentlyContinue
            $hint = "请确认 PowerPoint、Word、插件设置窗口和 URL 编辑窗口都已关闭，然后重试。"
            if ($webView2) {
                $hint = "检测到 msedgewebview2 仍在运行。请关闭相关 WebView2 应用窗口，或重启后再执行安装。"
            }

            throw "无法清理旧安装目录：${Path}`n${hint}`n原始错误：$($_.Exception.Message)"
        }
    }
}

function Get-PackageManifestEntries {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    $manifestPath = Join-Path $SourceRoot "PACKAGE.txt"
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        return @()
    }

    $entries = New-Object System.Collections.Generic.List[object]
    $inContents = $false
    foreach ($line in Get-Content -LiteralPath $manifestPath) {
        $lineText = if ($null -eq $line) { "" } else { [string]$line }
        $trimmed = $lineText.Trim()
        if (-not $inContents) {
            if ($trimmed -eq "Contents:") {
                $inContents = $true
            }

            continue
        }

        if ($trimmed -notmatch '^- (.+)$') {
            continue
        }

        $entry = $Matches[1].Trim()
        $isOptional = $entry -match '\s+\(if available\)\s*$'
        $relativePath = ($entry -replace '\s+\(if available\)\s*$', '').Trim()
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
            continue
        }

        $entries.Add([pscustomobject]@{
            RelativePath = $relativePath
            Optional = $isOptional
        }) | Out-Null
    }

    return $entries.ToArray()
}

function Assert-RelativePackagePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    if ([System.IO.Path]::IsPathRooted($RelativePath) -or
        $RelativePath -match '(^|[\\/])\.\.([\\/]|$)') {
        throw "Unsafe package manifest path: $RelativePath"
    }
}

function Copy-PackagePayloadToStaging {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,
        [Parameter(Mandatory = $true)]
        [string]$StagingRoot
    )

    if (Test-Path -LiteralPath $StagingRoot) {
        Remove-Item -LiteralPath $StagingRoot -Recurse -Force
    }

    New-Item -ItemType Directory -Force -Path $StagingRoot | Out-Null

    $manifestEntries = @(Get-PackageManifestEntries -SourceRoot $SourceRoot)
    if ($manifestEntries.Count -eq 0) {
        Copy-Item (Join-Path $SourceRoot "*") $StagingRoot -Recurse -Force
        return
    }

    $sourceRootFullPath = [System.IO.Path]::GetFullPath($SourceRoot)
    $sourcePrefix = $sourceRootFullPath.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
    $stagingFullPath = [System.IO.Path]::GetFullPath($StagingRoot)
    $stagingPrefix = $stagingFullPath.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar

    $copied = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($manifestEntry in $manifestEntries) {
        $relativePath = [string]$manifestEntry.RelativePath
        Assert-RelativePackagePath -RelativePath $relativePath

        $sourcePath = [System.IO.Path]::GetFullPath((Join-Path $sourceRootFullPath $relativePath))
        if (-not $sourcePath.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Package manifest path escapes source root: $relativePath"
        }

        if (-not (Test-Path -LiteralPath $sourcePath)) {
            if ([bool]$manifestEntry.Optional) {
                continue
            }

            throw "Package manifest file not found: $relativePath"
        }

        $targetPath = [System.IO.Path]::GetFullPath((Join-Path $stagingFullPath $relativePath))
        if (-not $targetPath.StartsWith($stagingPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Package manifest path escapes staging root: $relativePath"
        }

        if (-not $copied.Add($targetPath)) {
            continue
        }

        $targetDirectory = Split-Path -Parent $targetPath
        if (-not (Test-Path -LiteralPath $targetDirectory)) {
            New-Item -ItemType Directory -Force -Path $targetDirectory | Out-Null
        }

        Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    }
}

function Read-BuildInfoValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $candidatePaths = @(
        (Join-Path $Root "BuildInfo.txt"),
        (Join-Path $Root "PACKAGE.txt"),
        (Join-Path $Root "bin\\BuildInfo.txt")
    )

    foreach ($candidatePath in $candidatePaths) {
        if (-not (Test-Path -LiteralPath $candidatePath)) {
            continue
        }

        foreach ($line in Get-Content -LiteralPath $candidatePath) {
            if ($line -match ('^\s*' + [regex]::Escape($Name) + '\s*[:=]\s*(.*)\s*$')) {
                return $Matches[1].Trim()
            }
        }
    }

    return ""
}

try {
    if (-not (Test-Path (Join-Path $sourceBin "DrawioPpt.PowerPointAddIn.dll"))) {
        throw "Release package root not found. Expected: $sourceBin\\DrawioPpt.PowerPointAddIn.dll"
    }

    if (-not (Test-Path (Join-Path $sourceBin "DrawioPpt.WordAddIn.dll"))) {
        throw "Release package root not found. Expected: $sourceBin\\DrawioPpt.WordAddIn.dll"
    }

    $runningPowerPoint = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
    if ($runningPowerPoint) {
        throw "请先关闭正在运行的 PowerPoint，再执行安装。"
    }

    $runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($runningWord) {
        throw "请先关闭正在运行的 Word，再执行安装。"
    }

    Copy-PackagePayloadToStaging -SourceRoot $sourceRoot -StagingRoot $stagingRoot

    Remove-ExistingInstallRoot -Path $targetRoot

    New-Item -ItemType Directory -Force -Path $targetRoot | Out-Null
    Copy-Item (Join-Path $stagingRoot "*") $targetRoot -Recurse -Force

    $registerScript = Join-Path $targetRoot "scripts\\register-office-addins.ps1"
    if (-not (Test-Path $registerScript)) {
        throw "Register script not found after install: $registerScript"
    }

    & powershell.exe -ExecutionPolicy Bypass -File $registerScript -PowerPointAssemblyPath $targetAssemblyPath -WordAssemblyPath $targetWordAssemblyPath
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    $verifyScript = Join-Path $targetRoot "scripts\\verify-office-install.ps1"
    if (-not (Test-Path $verifyScript)) {
        throw "Install verification script not found after install: $verifyScript"
    }

    & powershell.exe -ExecutionPolicy Bypass -File $verifyScript -InstallRoot $targetRoot -SkipComLoad
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Write-Host "DrawioPpt installed successfully."
    Write-Host "  InstallRoot: $targetRoot"
    Write-Host "  PowerPoint:  $targetAssemblyPath"
    Write-Host "  Word:        $targetWordAssemblyPath"
    $installedBuildId = Read-BuildInfoValue -Root $targetRoot -Name "BuildId"
    if (-not [string]::IsNullOrWhiteSpace($installedBuildId)) {
        Write-Host "  BuildId:     $installedBuildId"
    }
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
