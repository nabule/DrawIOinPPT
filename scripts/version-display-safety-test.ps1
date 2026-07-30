$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

function Read-Text {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath
    )

    $path = Join-Path $repoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required file is missing: $RelativePath"
    }

    return Get-Content -LiteralPath $path -Raw
}

$buildInfoSource = Read-Text "src\DrawioPpt.Core\BuildInfo.cs"
if ($buildInfoSource -notmatch 'class\s+BuildInfo' -or
    $buildInfoSource -notmatch 'DisplayVersion' -or
    $buildInfoSource -notmatch 'BuildInfo\.txt' -or
    $buildInfoSource -notmatch 'PACKAGE\.txt') {
    throw "BuildInfo must expose DisplayVersion and read BuildInfo.txt/PACKAGE.txt."
}

$settingsFormSource = Read-Text "src\DrawioPpt.PowerPointAddIn\UI\SettingsForm.cs"
if ($settingsFormSource -notmatch 'BuildInfo\.DisplayVersion' -or
    $settingsFormSource -notmatch ([string]::Concat([char]0x7248, [char]0x672C, [char]0xff1a))) {
    throw "Shared SettingsForm must display the build version."
}

$pptRibbonXml = Read-Text "src\DrawioPpt.PowerPointAddIn\Ribbon\RibbonXmlProvider.cs"
$wordRibbonXml = Read-Text "src\DrawioPpt.WordAddIn\Ribbon\RibbonXmlProvider.cs"
if ($pptRibbonXml -notmatch 'Greensoft\.DrawioPpt\.Version' -or
    $pptRibbonXml -notmatch "getLabel='GetVersionSummary'") {
    throw "PowerPoint Ribbon must include a version label."
}

if ($wordRibbonXml -notmatch 'Greensoft\.DrawioWord\.Version' -or
    $wordRibbonXml -notmatch "getLabel='GetVersionSummary'") {
    throw "Word Ribbon must include a version label."
}

$pptRibbonController = Read-Text "src\DrawioPpt.PowerPointAddIn\Ribbon\RibbonController.cs"
$wordRibbonController = Read-Text "src\DrawioPpt.WordAddIn\Ribbon\RibbonController.cs"
if ($pptRibbonController -notmatch 'GetVersionSummary' -or
    $wordRibbonController -notmatch 'GetVersionSummary') {
    throw "Both Ribbon controllers must expose GetVersionSummary."
}

$pptHost = Read-Text "src\DrawioPpt.PowerPointAddIn\Services\AddInHost.cs"
$wordHost = Read-Text "src\DrawioPpt.WordAddIn\Services\AddInHost.cs"
if ($pptHost -notmatch 'BuildInfo\.DisplayVersion' -or
    $wordHost -notmatch 'BuildInfo\.DisplayVersion') {
    throw "Both hosts must return BuildInfo.DisplayVersion for the Ribbon."
}

$pptConnect = Read-Text "src\DrawioPpt.PowerPointAddIn\Connect.cs"
$wordConnect = Read-Text "src\DrawioPpt.WordAddIn\Connect.cs"
foreach ($connect in @(
    @{ Name = "PowerPoint"; Source = $pptConnect },
    @{ Name = "Word"; Source = $wordConnect }
)) {
    if ($connect.Source -notmatch 'public\s+string\s+GetVersionSummary\s*\(\s*IRibbonControl\s+control\s*\)') {
        throw "$($connect.Name) Connect must expose the GetVersionSummary Office Ribbon callback."
    }

    if ($connect.Source -notmatch '_ribbonController\.GetVersionSummary\(control\)') {
        throw "$($connect.Name) Connect must forward GetVersionSummary to its Ribbon controller."
    }
}

if ($wordConnect -notmatch 'interface\s+IWordAddInAutomation' -or
    $wordConnect -notmatch 'string\s+GetVersionSummary\s*\(\s*\)' -or
    $wordConnect -notmatch 'class\s+WordAddInAutomation') {
    throw "Word automation must expose GetVersionSummary for installed-host verification."
}

if ($pptConnect -notmatch 'interface\s+IPowerPointAddInAutomation' -or
    $pptConnect -notmatch 'class\s+PowerPointAddInAutomation' -or
    $pptConnect -notmatch '_comAddIn\.Object\s*=\s*_automation') {
    throw "PowerPoint automation must expose GetVersionSummary for installed-host verification."
}

$buildScript = Read-Text "scripts\build.ps1"
$packageScript = Read-Text "scripts\package-release.ps1"
$installScript = Read-Text "scripts\install-release.ps1"
$verifyInstallScript = Read-Text "scripts\verify-office-install.ps1"
$writerScript = Read-Text "scripts\write-build-info.ps1"
if ($buildScript -notmatch 'write-build-info\.ps1' -or
    $packageScript -notmatch 'write-build-info\.ps1') {
    throw "Build and package scripts must generate BuildInfo.txt."
}

if ($buildScript -match 'powershell\.exe[^\r\n]+write-build-info\.ps1' -or
    $packageScript -match 'powershell\.exe[^\r\n]+write-build-info\.ps1') {
    throw "Build info writer must be invoked in-process so all OutputRoots are preserved."
}

if ($writerScript -notmatch 'rev-parse' -or
    $writerScript -notmatch '--short' -or
    $writerScript -notmatch 'HEAD' -or
    $writerScript -notmatch 'BuildId' -or
    $writerScript -notmatch 'GitShortHash') {
    throw "Build info writer must record the Git short hash and BuildId."
}

if ($packageScript -notmatch 'BuildId:' -or
    $packageScript -notmatch 'GitShortHash:' -or
    $packageScript -notmatch 'BuildInfo\.txt') {
    throw "Package manifest must include BuildId/GitShortHash and ship BuildInfo.txt."
}

if ($installScript -notmatch 'BuildId' -or
    $installScript -notmatch 'BuildInfo\.txt') {
    throw "Install script must report the installed BuildId."
}

if ($verifyInstallScript -notmatch 'Read-OfficeAddInVersion' -or
    $verifyInstallScript -notmatch 'OfficeAddInVersionReader' -or
    $verifyInstallScript -notmatch 'Add-Type' -or
    $verifyInstallScript -notmatch 'InvokeMember' -or
    $verifyInstallScript -notmatch 'GetVersionSummary' -or
    $verifyInstallScript -notmatch 'COM version callback verified' -or
    $verifyInstallScript -notmatch 'BuildInfo\.txt') {
    throw "Installed Office verification must call the COM version callback and compare it to BuildInfo.txt."
}

Write-Host "VERSION_DISPLAY_SAFETY_TEST_PASS"
