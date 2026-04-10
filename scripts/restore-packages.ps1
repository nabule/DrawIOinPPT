param()

$ErrorActionPreference = "Stop"

$packageRoot = Join-Path $PSScriptRoot "..\\packages"
$webView2Version = "1.0.3856.49"
$packageId = "Microsoft.Web.WebView2"
$packageFolderName = "$packageId.$webView2Version"
$packageDirectory = Join-Path $packageRoot $packageFolderName
$packageMarker = Join-Path $packageDirectory "lib\\net462\\Microsoft.Web.WebView2.WinForms.dll"
$packageArchive = Join-Path $packageRoot ($packageFolderName + ".nupkg")
$packageUrl = "https://api.nuget.org/v3-flatcontainer/microsoft.web.webview2/$webView2Version/microsoft.web.webview2.$webView2Version.nupkg"

if (Test-Path $packageMarker) {
    Write-Host "NuGet package already restored: $packageFolderName"
    return
}

if (-not (Test-Path $packageRoot)) {
    New-Item -ItemType Directory -Path $packageRoot | Out-Null
}

if (-not (Test-Path $packageArchive)) {
    Write-Host "Downloading $packageFolderName..."
    Invoke-WebRequest -Uri $packageUrl -OutFile $packageArchive
}

if (Test-Path $packageDirectory) {
    Remove-Item $packageDirectory -Recurse -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($packageArchive, $packageDirectory)

Write-Host "Restored NuGet package: $packageFolderName"
