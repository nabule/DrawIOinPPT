$checks = @()

function Add-Check {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail
    )

    $script:checks += New-Object psobject -Property @{
        Name = $Name
        Passed = $Passed
        Detail = $Detail
    }
}

$msbuildPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\MSBuild.exe"
$regasmPath = "C:\\Windows\\Microsoft.NET\\Framework64\\v4.0.30319\\RegAsm.exe"
$powerPointExe = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
$officeInterop = "C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.PowerPoint\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.PowerPoint.dll"
$officeCore = "C:\\Windows\\assembly\\GAC_MSIL\\office\\15.0.0.0__71e9bce111e9429c\\OFFICE.DLL"
$extensibility = "C:\\Windows\\assembly\\GAC\\Extensibility\\7.0.3300.0__b03f5f7f11d50a3a\\extensibility.dll"

Add-Check -Name "MSBuild" -Passed (Test-Path $msbuildPath) -Detail $msbuildPath
Add-Check -Name "RegAsm" -Passed (Test-Path $regasmPath) -Detail $regasmPath
Add-Check -Name "PowerPoint" -Passed (Test-Path $powerPointExe) -Detail $powerPointExe
Add-Check -Name "Interop.PowerPoint" -Passed (Test-Path $officeInterop) -Detail $officeInterop
Add-Check -Name "Office Core" -Passed (Test-Path $officeCore) -Detail $officeCore
Add-Check -Name "Extensibility" -Passed (Test-Path $extensibility) -Detail $extensibility

$checks | Format-Table -AutoSize

$failed = $checks | Where-Object { -not $_.Passed }
if ($failed.Count -gt 0) {
    exit 1
}

