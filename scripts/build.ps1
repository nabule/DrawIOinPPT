param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64"
)

$solutionPath = Join-Path $PSScriptRoot "..\\DrawioPpt.sln"
$msbuildPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\MSBuild.exe"
$restoreScriptPath = Join-Path $PSScriptRoot "restore-packages.ps1"

if (-not (Test-Path $solutionPath)) {
    throw "Solution file not found: $solutionPath"
}

if (-not (Test-Path $msbuildPath)) {
    throw "MSBuild.exe not found: $msbuildPath"
}

if (-not (Test-Path $restoreScriptPath)) {
    throw "Package restore script not found: $restoreScriptPath"
}

& powershell -ExecutionPolicy Bypass -File $restoreScriptPath

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

& $msbuildPath $solutionPath /p:Configuration=$Configuration /p:Platform=$Platform

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
