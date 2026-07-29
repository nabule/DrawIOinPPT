param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64",
    [string]$Version = "v1.0.8"
)

$solutionPath = Join-Path $PSScriptRoot "..\\DrawioPpt.sln"
$msbuildPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\MSBuild.exe"
$restoreScriptPath = Join-Path $PSScriptRoot "restore-packages.ps1"
$buildInfoScriptPath = Join-Path $PSScriptRoot "write-build-info.ps1"

if (-not (Test-Path $solutionPath)) {
    throw "Solution file not found: $solutionPath"
}

if (-not (Test-Path $msbuildPath)) {
    throw "MSBuild.exe not found: $msbuildPath"
}

if (-not (Test-Path $restoreScriptPath)) {
    throw "Package restore script not found: $restoreScriptPath"
}

if (-not (Test-Path $buildInfoScriptPath)) {
    throw "Build info script not found: $buildInfoScriptPath"
}

& powershell.exe -ExecutionPolicy Bypass -File $restoreScriptPath

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

& $msbuildPath $solutionPath /p:Configuration=$Configuration /p:Platform=$Platform

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$outputRoots = @(
    (Join-Path $repoRoot ("src\\DrawioPpt.Core\\bin\\{0}\\{1}" -f $Platform, $Configuration)),
    (Join-Path $repoRoot ("src\\DrawioPpt.PowerPointAddIn\\bin\\{0}\\{1}" -f $Platform, $Configuration)),
    (Join-Path $repoRoot ("src\\DrawioPpt.WordAddIn\\bin\\{0}\\{1}" -f $Platform, $Configuration))
)

& $buildInfoScriptPath -Version $Version -OutputRoots $outputRoots
