param(
    [string]$Configuration = "Debug",
    [switch]$X86
)

$platformFolder = "x64"
if ($X86) {
    $platformFolder = "x86"
}

$dllPath = Join-Path $PSScriptRoot ("..\\src\\DrawioPpt.PowerPointAddIn\\bin\\{0}\\{1}\\DrawioPpt.PowerPointAddIn.dll" -f $platformFolder, $Configuration)
$dllPath = [System.IO.Path]::GetFullPath($dllPath)

if (-not (Test-Path $dllPath)) {
    throw "Add-in assembly not found: $dllPath"
}

Add-Type -AssemblyName "System.Runtime.InteropServices"

$assemblyName = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath)
$assemblyDisplayName = $assemblyName.FullName
$assemblyVersion = $assemblyName.Version.ToString()
$runtimeVersion = "v4.0.30319"
$codeBase = (New-Object System.Uri($dllPath)).AbsoluteUri

$progId = "Greensoft.DrawioPptAddIn"
$className = "DrawioPpt.PowerPointAddIn.Connect"
$clsid = "{0B8996D8-D6B9-4D61-8E8C-6F2081BFEA31}"
$implementedCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
$mscoree = "mscoree.dll"

function Remove-RegistryKeyTree {
    param(
        [string]$SubKey
    )

    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($SubKey, $false)
}

function Set-RegistryValue {
    param(
        [string]$SubKey,
        [AllowNull()][string]$Name,
        [AllowNull()]$Value,
        [Microsoft.Win32.RegistryValueKind]$Kind = [Microsoft.Win32.RegistryValueKind]::String
    )

    $key = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($SubKey)
    if ($null -eq $key) {
        throw "Unable to create registry key: HKCU\\$SubKey"
    }

    try {
        if ([string]::IsNullOrEmpty($Name)) {
            $key.SetValue("", $Value, $Kind)
        }
        else {
            $key.SetValue($Name, $Value, $Kind)
        }
    }
    finally {
        $key.Dispose()
    }
}

$legacyAddinKey = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\$className"
Remove-RegistryKeyTree -SubKey $legacyAddinKey
Remove-RegistryKeyTree -SubKey ("Software\\Microsoft\\Office\\PowerPoint\\Addins\\$progId")
Remove-RegistryKeyTree -SubKey ("Software\\Classes\\$progId")
Remove-RegistryKeyTree -SubKey ("Software\\Classes\\CLSID\\$clsid")

Set-RegistryValue -SubKey ("Software\\Classes\\$progId") -Name $null -Value $className
Set-RegistryValue -SubKey ("Software\\Classes\\$progId\\CLSID") -Name $null -Value $clsid

$clsidRoot = "Software\\Classes\\CLSID\\$clsid"
Set-RegistryValue -SubKey $clsidRoot -Name $null -Value $className
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32" -Name $null -Value $mscoree
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32" -Name "ThreadingModel" -Value "Both"
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32" -Name "Class" -Value $className
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32" -Name "Assembly" -Value $assemblyDisplayName
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32" -Name "RuntimeVersion" -Value $runtimeVersion
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32" -Name "CodeBase" -Value $codeBase
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32\\$assemblyVersion" -Name "Class" -Value $className
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32\\$assemblyVersion" -Name "Assembly" -Value $assemblyDisplayName
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32\\$assemblyVersion" -Name "RuntimeVersion" -Value $runtimeVersion
Set-RegistryValue -SubKey "$clsidRoot\\InprocServer32\\$assemblyVersion" -Name "CodeBase" -Value $codeBase
Set-RegistryValue -SubKey "$clsidRoot\\ProgId" -Name $null -Value $progId
Set-RegistryValue -SubKey "$clsidRoot\\Implemented Categories\\$implementedCategory" -Name $null -Value ""

$officeAddinKey = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\$progId"
Set-RegistryValue -SubKey $officeAddinKey -Name "FriendlyName" -Value "DrawioPpt"
Set-RegistryValue -SubKey $officeAddinKey -Name "Description" -Value "Edit Draw.io diagrams inside PowerPoint."
Set-RegistryValue -SubKey $officeAddinKey -Name "LoadBehavior" -Value 3 -Kind ([Microsoft.Win32.RegistryValueKind]::DWord)
Set-RegistryValue -SubKey $officeAddinKey -Name "CommandLineSafe" -Value 0 -Kind ([Microsoft.Win32.RegistryValueKind]::DWord)

Write-Host "Registered DrawioPpt for current user."
