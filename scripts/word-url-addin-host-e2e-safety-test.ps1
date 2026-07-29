$ErrorActionPreference = "Stop"

$sourcePath = Join-Path $PSScriptRoot "word-url-addin-host-e2e.ps1"
$source = Get-Content -LiteralPath $sourcePath -Raw

if ($source -match '\$serverPort\s*=\s*Get-Random') {
    throw "The real Word URL E2E still chooses a random port without reserving an available loopback port."
}

if ($source -match 'Get-NewWordAutomationProcessIds') {
    throw "The real Word URL E2E still identifies every Word process created after the test starts."
}

if ($source -notmatch 'TestWordProcessIdentity=' -or
    $source -notmatch 'GetWindowThreadProcessId' -or
    $source -notmatch 'StartTicks' -or
    $source -notmatch 'processIdentityPath' -or
    $source -notmatch 'Read-TestWordProcessIdentities' -or
    $source -notmatch 'File\.AppendAllText') {
    throw "The real Word URL E2E does not bind cleanup to a Word process identity reported by its helper."
}

if ($source -notmatch 'StartWithRetry' -or $source -match 'function Write-UrlTestSettings') {
    throw "The real Word URL E2E does not bind its mock server before writing the editor URL setting."
}

if ($source -notmatch 'foreach \(\$cleanupAction in \$cleanupActions\)' -or
    $source -notmatch 'Cleanup failure') {
    throw "The real Word URL E2E cleanup steps are not independently isolated and summarized."
}

if ($source -notmatch 'WordProcessNaturalExitObserved' -or
    $source -notmatch 'Wait-TestWordProcessExit' -or
    $source -notmatch 'TestWordProcessStopped=') {
    throw "The real Word URL E2E does not wait for test-owned Word to exit naturally before forced cleanup."
}

Write-Host "WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS"
