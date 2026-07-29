param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "Greensoft\\DrawioPpt")
)

$ErrorActionPreference = "Stop"

$sourceRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$sourceBin = Join-Path $sourceRoot "bin"
$targetRoot = [System.IO.Path]::GetFullPath($InstallRoot)
$targetAssemblyPath = Join-Path $targetRoot "bin\\DrawioPpt.PowerPointAddIn.dll"
$targetWordAssemblyPath = Join-Path $targetRoot "bin\\DrawioPpt.WordAddIn.dll"

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

Remove-ExistingInstallRoot -Path $targetRoot

New-Item -ItemType Directory -Force -Path $targetRoot | Out-Null
Copy-Item (Join-Path $sourceRoot "*") $targetRoot -Recurse -Force

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
