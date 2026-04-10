param(
    [string]$Configuration = "Debug",
    [switch]$X86
)

function Remove-RegistryKeyTree {
    param(
        [string]$SubKey
    )

    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($SubKey, $false)
}

$progId = "Greensoft.DrawioPptAddIn"
$className = "DrawioPpt.PowerPointAddIn.Connect"
$clsid = "{0B8996D8-D6B9-4D61-8E8C-6F2081BFEA31}"

Remove-RegistryKeyTree -SubKey ("Software\\Microsoft\\Office\\PowerPoint\\Addins\\$progId")
Remove-RegistryKeyTree -SubKey ("Software\\Microsoft\\Office\\PowerPoint\\Addins\\$className")
Remove-RegistryKeyTree -SubKey ("Software\\Classes\\$progId")
Remove-RegistryKeyTree -SubKey ("Software\\Classes\\CLSID\\$clsid")

Write-Host "Unregistered DrawioPpt for current user."
