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
$net48ReferenceAssemblies = "C:\\Program Files (x86)\\Reference Assemblies\\Microsoft\\Framework\\.NETFramework\\v4.8"
$restoreScript = Join-Path $PSScriptRoot "restore-packages.ps1"
$webView2RuntimeRegistryPath = "Registry::HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
$webView2RuntimeVersion = (Get-ItemProperty -Path $webView2RuntimeRegistryPath -Name pv -ErrorAction SilentlyContinue).pv
$webView2RuntimeDetail = $webView2RuntimeVersion
if ([string]::IsNullOrWhiteSpace($webView2RuntimeDetail)) {
    $webView2RuntimeDetail = $webView2RuntimeRegistryPath
}

Add-Check -Name "MSBuild" -Passed (Test-Path $msbuildPath) -Detail $msbuildPath
Add-Check -Name "RegAsm" -Passed (Test-Path $regasmPath) -Detail $regasmPath
Add-Check -Name "PowerPoint" -Passed (Test-Path $powerPointExe) -Detail $powerPointExe
Add-Check -Name "Interop.PowerPoint" -Passed (Test-Path $officeInterop) -Detail $officeInterop
Add-Check -Name "Office Core" -Passed (Test-Path $officeCore) -Detail $officeCore
Add-Check -Name "Extensibility" -Passed (Test-Path $extensibility) -Detail $extensibility
Add-Check -Name ".NET 4.8 Ref Pack" -Passed (Test-Path $net48ReferenceAssemblies) -Detail $net48ReferenceAssemblies
Add-Check -Name "Restore Script" -Passed (Test-Path $restoreScript) -Detail $restoreScript
Add-Check -Name "WebView2 Runtime" -Passed (-not [string]::IsNullOrWhiteSpace($webView2RuntimeVersion)) -Detail $webView2RuntimeDetail

$checks | Format-Table -AutoSize

$failed = $checks | Where-Object { -not $_.Passed }
if ($failed.Count -gt 0) {
    exit 1
}
