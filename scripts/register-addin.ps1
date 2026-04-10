param(
    [string]$Configuration = "Debug",
    [switch]$X86
)

$frameworkRoot = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319"
if ($X86) {
    $frameworkRoot = "C:\\Windows\\Microsoft.NET\\Framework\\v4.0.30319"
}

$regasmPath = Join-Path $frameworkRoot "RegAsm.exe"
$platformFolder = "x64"
if ($X86) {
    $platformFolder = "x86"
}

$dllPath = Join-Path $PSScriptRoot ("..\\src\\DrawioPpt.PowerPointAddIn\\bin\\{0}\\{1}\\DrawioPpt.PowerPointAddIn.dll" -f $platformFolder, $Configuration)

if (-not (Test-Path $regasmPath)) {
    throw "RegAsm.exe not found: $regasmPath"
}

if (-not (Test-Path $dllPath)) {
    throw "Add-in assembly not found: $dllPath"
}

& $regasmPath $dllPath /codebase

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
