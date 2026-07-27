param(
    [string]$SourcePath
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Join-Path $PSScriptRoot "..\src\DrawioPpt.PowerPointAddIn\UI\UrlDiagramEditorForm.cs"
}

if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "URL editor source file was not found: $SourcePath"
}

$source = Get-Content -LiteralPath $SourcePath -Raw

if ($source -match 'EnsureCoreWebView2Async\s*\(\s*\)') {
    throw "URL editor still initializes WebView2 with the host default data folder."
}

if ($source -notmatch 'CoreWebView2Environment\.CreateAsync\s*\(\s*null\s*,\s*GetWebView2UserDataFolder\s*\(\s*\)\s*\)') {
    throw "URL editor does not create a WebView2 environment with its per-user data folder."
}

if ($source -notmatch 'EnsureCoreWebView2Async\s*\(\s*environment\s*\)') {
    throw "URL editor does not initialize WebView2 with the explicit environment."
}

if ($source -notmatch 'Environment\.SpecialFolder\.LocalApplicationData' -or
    $source -notmatch '"Greensoft"' -or
    $source -notmatch '"DrawioPpt"' -or
    $source -notmatch '"WebView2"') {
    throw "URL editor data folder is not located under the current user's LocalApplicationData directory."
}

if ($source -notmatch 'Directory\.CreateDirectory\s*\(\s*userDataFolder\s*\)') {
    throw "URL editor does not ensure that the per-user WebView2 data folder exists."
}

Write-Host "WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS"
