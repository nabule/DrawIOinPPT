param(
    [string]$Configuration = "Debug"
)

$solutionPath = Join-Path $PSScriptRoot "..\\DrawioPpt.sln"
$msbuildPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\MSBuild.exe"

if (-not (Test-Path $solutionPath)) {
    throw "Solution file not found: $solutionPath"
}

if (-not (Test-Path $msbuildPath)) {
    throw "MSBuild.exe not found: $msbuildPath"
}

& $msbuildPath $solutionPath /p:Configuration=$Configuration /p:Platform="Any CPU"

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

