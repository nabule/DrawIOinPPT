$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$installScriptUnderTest = Join-Path $repoRoot "scripts\install-release.ps1"

function New-PackageFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot
    )

    New-Item -ItemType Directory -Force -Path (Join-Path $PackageRoot "bin") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $PackageRoot "scripts") | Out-Null

    Copy-Item -LiteralPath $installScriptUnderTest -Destination (Join-Path $PackageRoot "scripts\install-release.ps1") -Force
    Set-Content -LiteralPath (Join-Path $PackageRoot "README.md") -Value "fixture package" -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $PackageRoot "BuildInfo.txt") -Value @(
        "ProductVersion: v-test",
        "BuildId: v-test+upgrade",
        "GitShortHash: upgrade"
    ) -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $PackageRoot "bin\DrawioPpt.PowerPointAddIn.dll") -Value "ppt-dll" -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $PackageRoot "bin\DrawioPpt.WordAddIn.dll") -Value "word-dll" -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $PackageRoot "scripts\register-office-addins.ps1") -Value @'
param(
    [string]$PowerPointAssemblyPath,
    [string]$WordAssemblyPath
)

if (-not (Test-Path -LiteralPath $PowerPointAssemblyPath)) {
    throw "PowerPoint assembly missing: $PowerPointAssemblyPath"
}

if (-not (Test-Path -LiteralPath $WordAssemblyPath)) {
    throw "Word assembly missing: $WordAssemblyPath"
}

Set-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) "register-called.txt") -Value "registered" -Encoding UTF8
'@ -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $PackageRoot "scripts\verify-office-install.ps1") -Value @'
param(
    [string]$InstallRoot,
    [switch]$SkipComLoad
)

if (-not (Test-Path -LiteralPath (Join-Path $InstallRoot "bin\DrawioPpt.PowerPointAddIn.dll"))) {
    throw "Installed PowerPoint assembly missing."
}

if (-not (Test-Path -LiteralPath (Join-Path $InstallRoot "bin\DrawioPpt.WordAddIn.dll"))) {
    throw "Installed Word assembly missing."
}

Set-Content -LiteralPath (Join-Path $InstallRoot "verify-called.txt") -Value "verified" -Encoding UTF8
'@ -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $PackageRoot "PACKAGE.txt") -Value @'
Package: DrawioPpt
Version: v-test
Contents:
- README.md
- PACKAGE.txt
- BuildInfo.txt
- bin\DrawioPpt.PowerPointAddIn.dll
- bin\DrawioPpt.WordAddIn.dll
- scripts\install-release.ps1
- scripts\register-office-addins.ps1
- scripts\verify-office-install.ps1
'@ -Encoding UTF8
}

function Invoke-InstallAndAssert {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageRoot,
        [Parameter(Mandatory = $true)]
        [string]$InstallRoot,
        [Parameter(Mandatory = $true)]
        [string]$ScenarioName
    )

    Set-Content -LiteralPath (Join-Path $InstallRoot "stale-old-file.txt") -Value "stale" -Encoding UTF8

    $installScript = Join-Path $PackageRoot "scripts\install-release.ps1"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installScript -InstallRoot $InstallRoot
    if ($LASTEXITCODE -ne 0) {
        throw "$ScenarioName install failed with exit code $LASTEXITCODE."
    }

    $requiredInstalledFiles = @(
        "README.md",
        "BuildInfo.txt",
        "PACKAGE.txt",
        "bin\DrawioPpt.PowerPointAddIn.dll",
        "bin\DrawioPpt.WordAddIn.dll",
        "scripts\install-release.ps1",
        "scripts\register-office-addins.ps1",
        "scripts\verify-office-install.ps1",
        "register-called.txt",
        "verify-called.txt"
    )

    foreach ($relativePath in $requiredInstalledFiles) {
        $path = Join-Path $InstallRoot $relativePath
        if (-not (Test-Path -LiteralPath $path)) {
            throw "$ScenarioName did not install required file: $relativePath"
        }
    }

    if (Test-Path -LiteralPath (Join-Path $InstallRoot "stale-old-file.txt")) {
        throw "$ScenarioName kept a stale file from the previous install."
    }

    Write-Host "$ScenarioName=PASS"
}

$testRoot = Join-Path $env:TEMP ("DrawioPpt\upgrade-safety-" + [Guid]::NewGuid().ToString("N"))
try {
    $sameRoot = Join-Path $testRoot "same-root"
    New-PackageFixture -PackageRoot $sameRoot
    Invoke-InstallAndAssert -PackageRoot $sameRoot -InstallRoot $sameRoot -ScenarioName "SameRootUpgrade"

    $nestedInstallRoot = Join-Path $testRoot "nested-root"
    $nestedPackageRoot = Join-Path $nestedInstallRoot "upgrade-package"
    New-Item -ItemType Directory -Force -Path $nestedInstallRoot | Out-Null
    New-PackageFixture -PackageRoot $nestedPackageRoot
    Invoke-InstallAndAssert -PackageRoot $nestedPackageRoot -InstallRoot $nestedInstallRoot -ScenarioName "NestedPackageUpgrade"
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

Write-Host "INSTALL_RELEASE_UPGRADE_SAFETY_TEST_PASS"
