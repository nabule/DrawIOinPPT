param(
    [string]$Configuration = "Debug",
    [switch]$X86,
    [string]$PowerPointAssemblyPath,
    [string]$WordAssemblyPath
)

$ErrorActionPreference = "Stop"

function Invoke-RegistrationScript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [string]$Configuration,
        [switch]$X86,
        [string]$AssemblyPath
    )

    if (-not (Test-Path $ScriptPath)) {
        throw "Registration script not found: $ScriptPath"
    }

    $arguments = @("-ExecutionPolicy", "Bypass", "-File", $ScriptPath, "-Configuration", $Configuration)
    if ($X86) {
        $arguments += "-X86"
    }

    if (-not [string]::IsNullOrWhiteSpace($AssemblyPath)) {
        $arguments += @("-AssemblyPath", $AssemblyPath)
    }

    & powershell.exe @arguments
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

$powerPointRegisterScript = Join-Path $PSScriptRoot "register-addin.ps1"
$wordRegisterScript = Join-Path $PSScriptRoot "register-word-addin.ps1"

Invoke-RegistrationScript `
    -ScriptPath $powerPointRegisterScript `
    -Configuration $Configuration `
    -X86:$X86 `
    -AssemblyPath $PowerPointAssemblyPath

Invoke-RegistrationScript `
    -ScriptPath $wordRegisterScript `
    -Configuration $Configuration `
    -X86:$X86 `
    -AssemblyPath $WordAssemblyPath

Write-Host "Registered DrawioPpt Office add-ins for current user."
