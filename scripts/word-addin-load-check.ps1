param(
    [string]$Configuration = "Debug",
    [string]$AssemblyPath,
    [switch]$KeepRegistered
)

$ErrorActionPreference = "Stop"

$registerScript = Join-Path $PSScriptRoot "register-word-addin.ps1"
$unregisterScript = Join-Path $PSScriptRoot "unregister-word-addin.ps1"
$wordClsidKey = "HKCU:\Software\Classes\CLSID\{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}\InprocServer32"

function Assert-NoRunningWord {
    $runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($runningWord) {
        throw "Close Word before running the Word add-in load check."
    }
}

function Get-RegisteredWordAssemblyPath {
    if (-not (Test-Path $wordClsidKey)) {
        return $null
    }

    $codeBase = (Get-ItemProperty -Path $wordClsidKey -ErrorAction Stop).CodeBase
    if ([string]::IsNullOrWhiteSpace($codeBase)) {
        return $null
    }

    return ([System.Uri]$codeBase).LocalPath
}

function Register-WordAddIn {
    param(
        [string]$Configuration,
        [string]$AssemblyPath
    )

    $arguments = @("-ExecutionPolicy", "Bypass", "-File", $registerScript, "-Configuration", $Configuration)
    if (-not [string]::IsNullOrWhiteSpace($AssemblyPath)) {
        $arguments += @("-AssemblyPath", $AssemblyPath)
    }

    & powershell.exe @arguments
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

Assert-NoRunningWord

$previousAssemblyPath = Get-RegisteredWordAssemblyPath
$hadPreviousRegistration = -not [string]::IsNullOrWhiteSpace($previousAssemblyPath)

Register-WordAddIn -Configuration $Configuration -AssemblyPath $AssemblyPath

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $addin = $word.COMAddIns.Item("Greensoft.DrawioWordAddIn")
    $addin.Connect = $true
    Write-Host ("WordAddInConnect=" + $addin.Connect)
    if (-not [bool]$addin.Connect) {
        exit 1
    }
}
finally {
    if ($word -ne $null) {
        try {
            $saveChanges = [Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges
            $word.Quit([ref]$saveChanges) | Out-Null
        }
        catch {
        }
    }

    if (-not $KeepRegistered) {
        if ($hadPreviousRegistration) {
            Register-WordAddIn -Configuration $Configuration -AssemblyPath $previousAssemblyPath | Out-Null
        }
        else {
            & powershell.exe -ExecutionPolicy Bypass -File $unregisterScript -Configuration $Configuration | Out-Null
        }
    }
}
