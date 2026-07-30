param(
    [string]$Version = "v1.0.9",
    [string[]]$OutputRoots = @()
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

function Invoke-GitText {
    param(
        [string[]]$Arguments
    )

    try {
        $output = & git -C $repoRoot @Arguments 2>$null
        if ($LASTEXITCODE -ne 0) {
            return ""
        }

        return ($output | Select-Object -First 1).ToString().Trim()
    }
    catch {
        return ""
    }
}

$gitShortHash = Invoke-GitText @("rev-parse", "--short", "HEAD")
if ([string]::IsNullOrWhiteSpace($gitShortHash)) {
    $gitShortHash = "unknown"
}

$trackedStatus = & git -C $repoRoot status --porcelain --untracked-files=no 2>$null
$gitIsDirty = ($LASTEXITCODE -eq 0 -and @($trackedStatus).Count -gt 0)
$dirtySuffix = if ($gitIsDirty) { "-dirty" } else { "" }
$buildId = if ($gitShortHash -eq "unknown") {
    $Version
} else {
    $Version + "+" + $gitShortHash + $dirtySuffix
}

$content = @(
    "ProductVersion: $Version",
    "BuildId: $buildId",
    "GitShortHash: $gitShortHash",
    "GitIsDirty: $gitIsDirty",
    "BuiltAtUtc: $([DateTime]::UtcNow.ToString("o"))"
)

foreach ($outputRoot in $OutputRoots) {
    if ([string]::IsNullOrWhiteSpace($outputRoot)) {
        continue
    }

    $fullRoot = [System.IO.Path]::GetFullPath($outputRoot)
    if (-not (Test-Path -LiteralPath $fullRoot)) {
        New-Item -ItemType Directory -Force -Path $fullRoot | Out-Null
    }

    Set-Content -LiteralPath (Join-Path $fullRoot "BuildInfo.txt") -Value $content -Encoding UTF8
}

Write-Host "BuildId=$buildId"
Write-Host "GitShortHash=$gitShortHash"
