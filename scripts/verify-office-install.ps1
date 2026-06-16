param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "Greensoft\DrawioPpt"),
    [switch]$SkipComLoad
)

$ErrorActionPreference = "Stop"

$targetRoot = [System.IO.Path]::GetFullPath($InstallRoot)
$powerPointAssemblyPath = Join-Path $targetRoot "bin\DrawioPpt.PowerPointAddIn.dll"
$wordAssemblyPath = Join-Path $targetRoot "bin\DrawioPpt.WordAddIn.dll"

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        throw "Required installed file not found: $Path"
    }
}

function Get-CodeBaseUri {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return (New-Object System.Uri(([System.IO.Path]::GetFullPath($Path)))).AbsoluteUri
}

function Get-RegistryValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath,
        [AllowNull()]
        [string]$Name
    )

    if (-not (Test-Path $KeyPath)) {
        throw "Registry key not found: $KeyPath"
    }

    $key = Get-Item -Path $KeyPath
    try {
        return $key.GetValue($Name)
    }
    finally {
        $key.Close()
    }
}

function Assert-RegistryValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath,
        [AllowNull()]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Expected
    )

    $actual = [string](Get-RegistryValue -KeyPath $KeyPath -Name $Name)
    if ($actual -ne $Expected) {
        $displayName = if ([string]::IsNullOrEmpty($Name)) { "(default)" } else { $Name }
        throw "Unexpected registry value at ${KeyPath}\${displayName}. Expected '$Expected', got '$actual'."
    }
}

function Assert-RegistryDWord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyPath,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [int]$Expected
    )

    $actual = [int](Get-RegistryValue -KeyPath $KeyPath -Name $Name)
    if ($actual -ne $Expected) {
        throw "Unexpected registry value at ${KeyPath}\${Name}. Expected $Expected, got $actual."
    }
}

function Assert-AddInRegistration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName,
        [Parameter(Mandatory = $true)]
        [string]$ProgId,
        [Parameter(Mandatory = $true)]
        [string]$ClassName,
        [Parameter(Mandatory = $true)]
        [string]$Clsid,
        [Parameter(Mandatory = $true)]
        [string]$AssemblyPath
    )

    $officeAddinKey = "HKCU:\Software\Microsoft\Office\$HostName\Addins\$ProgId"
    $progIdKey = "HKCU:\Software\Classes\$ProgId"
    $progIdClsidKey = "HKCU:\Software\Classes\$ProgId\CLSID"
    $clsidKey = "HKCU:\Software\Classes\CLSID\$Clsid"
    $inprocKey = "$clsidKey\InprocServer32"
    $expectedCodeBase = Get-CodeBaseUri -Path $AssemblyPath

    Assert-RegistryDWord -KeyPath $officeAddinKey -Name "LoadBehavior" -Expected 3
    Assert-RegistryDWord -KeyPath $officeAddinKey -Name "CommandLineSafe" -Expected 0
    Assert-RegistryValue -KeyPath $progIdKey -Name $null -Expected $ClassName
    Assert-RegistryValue -KeyPath $progIdClsidKey -Name $null -Expected $Clsid
    Assert-RegistryValue -KeyPath $clsidKey -Name $null -Expected $ClassName
    Assert-RegistryValue -KeyPath $inprocKey -Name $null -Expected "mscoree.dll"
    Assert-RegistryValue -KeyPath $inprocKey -Name "Class" -Expected $ClassName
    Assert-RegistryValue -KeyPath $inprocKey -Name "CodeBase" -Expected $expectedCodeBase

    Write-Host "$HostName registration verified: $ProgId"
}

function Assert-NotDisabledByOffice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName,
        [Parameter(Mandatory = $true)]
        [string[]]$Patterns
    )

    foreach ($version in @("16.0", "15.0")) {
        $disabledKey = "HKCU:\Software\Microsoft\Office\$version\$HostName\Resiliency\DisabledItems"
        if (-not (Test-Path $disabledKey)) {
            continue
        }

        $properties = Get-ItemProperty -Path $disabledKey
        foreach ($property in $properties.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" }) {
            $bytes = [byte[]]$property.Value
            $unicode = [System.Text.Encoding]::Unicode.GetString($bytes)
            $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)
            foreach ($pattern in $Patterns) {
                if ($unicode.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                    $ascii.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    throw "$HostName has a disabled Office add-in entry matching '$pattern' at $disabledKey value '$($property.Name)'. Re-enable or clear the disabled item before using the add-in."
                }
            }
        }
    }
}

function Assert-NoRunningOfficeProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessName
    )

    $running = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($running) {
        throw "Please close running $ProcessName before COM load verification."
    }
}

function Assert-ComAddInConnects {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ApplicationProgId,
        [Parameter(Mandatory = $true)]
        [string]$ProcessName,
        [Parameter(Mandatory = $true)]
        [string]$ProgId
    )

    Assert-NoRunningOfficeProcess -ProcessName $ProcessName

    $application = $null
    try {
        $application = New-Object -ComObject $ApplicationProgId
        if ($ApplicationProgId -eq "Word.Application") {
            $application.Visible = $false
        }

        $addin = $application.COMAddIns.Item($ProgId)
        $addin.Connect = $true
        if (-not [bool]$addin.Connect) {
            throw "$ProgId did not connect in $ApplicationProgId."
        }

        Write-Host "$ApplicationProgId COM load verified: $ProgId"
    }
    finally {
        if ($application -ne $null) {
            try {
                if ($ApplicationProgId -eq "Word.Application") {
                    $saveChanges = 0
                    $application.Quit([ref]$saveChanges) | Out-Null
                }
                else {
                    $application.Quit() | Out-Null
                }
            }
            catch {
            }

            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($application)
        }
    }
}

Assert-FileExists -Path $powerPointAssemblyPath
Assert-FileExists -Path $wordAssemblyPath
Assert-FileExists -Path (Join-Path $targetRoot "scripts\register-office-addins.ps1")
Assert-FileExists -Path (Join-Path $targetRoot "scripts\register-word-addin.ps1")

Assert-AddInRegistration `
    -HostName "PowerPoint" `
    -ProgId "Greensoft.DrawioPptAddIn" `
    -ClassName "DrawioPpt.PowerPointAddIn.Connect" `
    -Clsid "{0B8996D8-D6B9-4D61-8E8C-6F2081BFEA31}" `
    -AssemblyPath $powerPointAssemblyPath

Assert-AddInRegistration `
    -HostName "Word" `
    -ProgId "Greensoft.DrawioWordAddIn" `
    -ClassName "DrawioPpt.WordAddIn.Connect" `
    -Clsid "{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}" `
    -AssemblyPath $wordAssemblyPath

Assert-NotDisabledByOffice `
    -HostName "Word" `
    -Patterns @("Greensoft.DrawioWordAddIn", "DrawioPpt.WordAddIn", "{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}", "DrawioWord")

Assert-NotDisabledByOffice `
    -HostName "PowerPoint" `
    -Patterns @("Greensoft.DrawioPptAddIn", "DrawioPpt.PowerPointAddIn", "{0B8996D8-D6B9-4D61-8E8C-6F2081BFEA31}", "DrawioPpt")

if (-not $SkipComLoad) {
    Assert-ComAddInConnects -ApplicationProgId "Word.Application" -ProcessName "WINWORD" -ProgId "Greensoft.DrawioWordAddIn"
    Assert-ComAddInConnects -ApplicationProgId "PowerPoint.Application" -ProcessName "POWERPNT" -ProgId "Greensoft.DrawioPptAddIn"
}

Write-Host "DrawioPpt Office installation verified."
