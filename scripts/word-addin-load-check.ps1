param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$registerScript = Join-Path $PSScriptRoot "register-word-addin.ps1"
$unregisterScript = Join-Path $PSScriptRoot "unregister-word-addin.ps1"

function Assert-NoRunningWord {
    $runningWord = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($runningWord) {
        throw "请先关闭正在运行的 Word，再执行 Word Add-in 加载检查。"
    }
}

Assert-NoRunningWord

& powershell.exe -ExecutionPolicy Bypass -File $registerScript -Configuration $Configuration
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

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

    & powershell.exe -ExecutionPolicy Bypass -File $unregisterScript -Configuration $Configuration | Out-Null
}
