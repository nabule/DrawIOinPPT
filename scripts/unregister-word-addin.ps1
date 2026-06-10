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

$progId = "Greensoft.DrawioWordAddIn"
$className = "DrawioPpt.WordAddIn.Connect"
$clsid = "{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}"

Remove-RegistryKeyTree -SubKey ("Software\\Microsoft\\Office\\Word\\Addins\\$progId")
Remove-RegistryKeyTree -SubKey ("Software\\Microsoft\\Office\\Word\\Addins\\$className")
Remove-RegistryKeyTree -SubKey ("Software\\Classes\\$progId")
Remove-RegistryKeyTree -SubKey ("Software\\Classes\\CLSID\\$clsid")

Write-Host "Unregistered DrawioWord for current user."
