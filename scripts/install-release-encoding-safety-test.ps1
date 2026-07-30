param(
    [string]$PackageRoot
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

function Assert-AsciiPowerShellFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,
        [Parameter(Mandatory = $true)]
        [string]$ScopeName
    )

    $scriptFiles = @(Get-ChildItem -LiteralPath $Root -Filter "*.ps1" -File -Recurse)
    if ($scriptFiles.Count -eq 0) {
        throw "$ScopeName did not contain any PowerShell scripts."
    }

    foreach ($scriptFile in $scriptFiles) {
        $bytes = [System.IO.File]::ReadAllBytes($scriptFile.FullName)
        $nonAsciiByte = @($bytes | Where-Object { $_ -gt 0x7f } | Select-Object -First 1)
        if ($nonAsciiByte.Count -ne 0) {
            throw "$ScopeName contains a non-ASCII PowerShell script: $($scriptFile.FullName)"
        }
    }
}

function Assert-InstallCommandTemplate {
    $packageScriptPath = Join-Path $repoRoot "scripts\package-release.ps1"
    $packageScriptText = Get-Content -LiteralPath $packageScriptPath -Raw
    $requiredFragments = @(
        'set "exitCode=%ERRORLEVEL%"',
        'DrawioPpt installation failed with exit code',
        'exit /b %exitCode%'
    )

    foreach ($fragment in $requiredFragments) {
        if (-not $packageScriptText.Contains($fragment)) {
            throw "Package install.cmd template is missing: $fragment"
        }
    }
}

function Assert-PackagedInstallCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $installCommandPath = Join-Path $Root "install.cmd"
    if (-not (Test-Path -LiteralPath $installCommandPath)) {
        throw "Packaged install.cmd was not found: $installCommandPath"
    }

    $installCommandText = Get-Content -LiteralPath $installCommandPath -Raw
    $requiredFragments = @(
        'set "exitCode=%ERRORLEVEL%"',
        'DrawioPpt installation failed with exit code',
        'exit /b %exitCode%'
    )

    foreach ($fragment in $requiredFragments) {
        if (-not $installCommandText.Contains($fragment)) {
            throw "Packaged install.cmd is missing: $fragment"
        }
    }

    Assert-AsciiPowerShellFiles -Root (Join-Path $Root "scripts") -ScopeName "Release package"
}

function Assert-InstallCommandFailurePropagation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $sourceInstallCommand = Join-Path $Root "install.cmd"
    $probeRoot = Join-Path $env:TEMP ("DrawioPpt\install-command-failure-" + [Guid]::NewGuid().ToString("N"))
    try {
        New-Item -ItemType Directory -Force -Path (Join-Path $probeRoot "scripts") | Out-Null
        Copy-Item -LiteralPath $sourceInstallCommand -Destination (Join-Path $probeRoot "install.cmd") -Force
        @'
Write-Error "INTENTIONAL_INSTALL_FAILURE"
exit 37
'@ | Set-Content -LiteralPath (Join-Path $probeRoot "scripts\install-release.ps1") -Encoding ASCII

        $installCommandPath = Join-Path $probeRoot "install.cmd"
        $originalErrorActionPreference = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            $installOutput = & cmd.exe /d /c ('"{0}"' -f $installCommandPath) 2>&1
            $installExitCode = [int]$LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $originalErrorActionPreference
        }
        $installOutputText = (@($installOutput) | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine

        if ($installExitCode -ne 37) {
            throw "install.cmd returned $installExitCode instead of the PowerShell failure exit code 37."
        }

        foreach ($expectedText in @(
                "INTENTIONAL_INSTALL_FAILURE",
                "DrawioPpt installation failed with exit code 37.")) {
            if (-not $installOutputText.Contains($expectedText)) {
                throw "install.cmd failure output is missing: $expectedText"
            }
        }
    }
    finally {
        if (Test-Path -LiteralPath $probeRoot) {
            Remove-Item -LiteralPath $probeRoot -Recurse -Force
        }
    }
}

Assert-AsciiPowerShellFiles -Root (Join-Path $repoRoot "scripts") -ScopeName "Repository"
$packageScriptPath = Join-Path $repoRoot "scripts\package-release.ps1"
if (Test-Path -LiteralPath $packageScriptPath) {
    Assert-InstallCommandTemplate
}

if (-not [string]::IsNullOrWhiteSpace($PackageRoot)) {
    $resolvedPackageRoot = [System.IO.Path]::GetFullPath($PackageRoot)
    Assert-PackagedInstallCommand -Root $resolvedPackageRoot
    Assert-InstallCommandFailurePropagation -Root $resolvedPackageRoot
}

Write-Host "INSTALL_RELEASE_ENCODING_SAFETY_TEST_PASS"
