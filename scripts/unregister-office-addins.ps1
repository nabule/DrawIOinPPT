param(
    [string]$Configuration = "Debug",
    [switch]$X86
)

$ErrorActionPreference = "Stop"

function Invoke-UnregistrationScript {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        [string]$Configuration,
        [switch]$X86
    )

    if (-not (Test-Path $ScriptPath)) {
        throw "Unregistration script not found: $ScriptPath"
    }

    $arguments = @("-ExecutionPolicy", "Bypass", "-File", $ScriptPath, "-Configuration", $Configuration)
    if ($X86) {
        $arguments += "-X86"
    }

    & powershell.exe @arguments
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

$powerPointUnregisterScript = Join-Path $PSScriptRoot "unregister-addin.ps1"
$wordUnregisterScript = Join-Path $PSScriptRoot "unregister-word-addin.ps1"

Invoke-UnregistrationScript `
    -ScriptPath $powerPointUnregisterScript `
    -Configuration $Configuration `
    -X86:$X86

Invoke-UnregistrationScript `
    -ScriptPath $wordUnregisterScript `
    -Configuration $Configuration `
    -X86:$X86

Write-Host "Unregistered DrawioPpt Office add-ins for current user."
