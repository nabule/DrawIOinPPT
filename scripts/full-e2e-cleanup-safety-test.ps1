param()

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$fullE2EPath = Join-Path $PSScriptRoot "full-e2e-test.ps1"
$complexMetadataPath = Join-Path $PSScriptRoot "word-complex-metadata-e2e.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\full-e2e-cleanup-safety-" + [Guid]::NewGuid().ToString("N"))
$dummyProcess = $null

function Import-FunctionDefinition {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Language.Ast]$Ast,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $definition = $Ast.Find(
        {
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                [string]::Equals($node.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)
        },
        $true)

    if ($null -eq $definition) {
        throw "Required function was not found in full-e2e-test.ps1: $Name"
    }

    $bodyText = $definition.Body.Extent.Text
    $bodyText = $bodyText.Substring(1, $bodyText.Length - 2)
    Set-Item `
        -Path ("Function:\script:" + $Name) `
        -Value ([scriptblock]::Create($bodyText))
}

function Get-FunctionDefinitionText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Language.Ast]$Ast,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $definition = $Ast.Find(
        {
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                [string]::Equals($node.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)
        },
        $true)

    if ($null -eq $definition) {
        throw "Required function was not found in full-e2e-test.ps1: $Name"
    }

    return $definition.Extent.Text
}

try {
    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $fullE2EPath,
        [ref]$tokens,
        [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw ("full-e2e-test.ps1 parse failed: " + ($parseErrors.Message -join "; "))
    }

    $exitStatements = @($ast.FindAll(
        {
            param($node)
            $node -is [System.Management.Automation.Language.ExitStatementAst]
        },
        $true))
    if ($exitStatements.Count -ne 1 -or
        $exitStatements[0].Extent.Text.Trim() -ne "exit 1") {
        throw "Full E2E must have exactly one final exit statement, 'exit 1'; found $($exitStatements.Count)."
    }

    $fullE2ESource = Get-Content -Raw -LiteralPath $fullE2EPath
    if ($fullE2ESource -match 'exit\s+\$LASTEXITCODE') {
        throw "Full E2E still contains an early exit based on LASTEXITCODE."
    }

    Import-FunctionDefinition -Ast $ast -Name "Stop-TestOwnedProcess"
    Import-FunctionDefinition -Ast $ast -Name "Invoke-CheckedPowerShellScript"
    Import-FunctionDefinition -Ast $ast -Name "Get-ErrorDetail"
    Import-FunctionDefinition -Ast $ast -Name "Get-FullE2EFailureDetails"
    Import-FunctionDefinition -Ast $ast -Name "Add-FullE2EFailureResult"

    $complexSource = Get-Content -Raw -LiteralPath $complexMetadataPath
    if ($complexSource -notmatch '\[string\]\$ProcessIdentityPath' -or
        $complexSource -notmatch 'GetWindowThreadProcessId' -or
        $complexSource -notmatch 'TestWordProcessIdentity=') {
        throw "Complex metadata E2E does not persist the owned Word PID and start-time identity."
    }

    if ($fullE2ESource -notmatch '-ProcessIdentityPath\s+\$complexWordIdentityPath') {
        throw "Full E2E does not pass its unique Word identity path to the installed complex metadata test."
    }

    $dummyProcess = Start-Process `
        -FilePath "powershell.exe" `
        -ArgumentList @("-NoProfile", "-Command", "Start-Sleep -Seconds 120") `
        -WindowStyle Hidden `
        -PassThru
    $dummyProcess.Refresh()
    $startTicks = $dummyProcess.StartTime.ToUniversalTime().Ticks
    $wrongIdentityPath = Join-Path $tempRoot "wrong-start.txt"
    $correctIdentityPath = Join-Path $tempRoot "correct-start.txt"
    Set-Content -LiteralPath $wrongIdentityPath -Value ("TestWordProcessIdentity={0}:{1}" -f $dummyProcess.Id, ($startTicks + 1)) -Encoding UTF8
    Set-Content -LiteralPath $correctIdentityPath -Value ("TestWordProcessIdentity={0}:{1}" -f $dummyProcess.Id, $startTicks) -Encoding UTF8

    Stop-TestOwnedProcess `
        -IdentityPath $wrongIdentityPath `
        -ExpectedProcessName "powershell" `
        -TimeoutSeconds 5
    if ($dummyProcess.HasExited) {
        throw "A process with a mismatched start time was terminated."
    }

    Stop-TestOwnedProcess `
        -IdentityPath $correctIdentityPath `
        -ExpectedProcessName "powershell" `
        -TimeoutSeconds 5
    $dummyProcess.WaitForExit(5000) | Out-Null
    if (-not $dummyProcess.HasExited) {
        throw "The matching test-owned process was not terminated."
    }

    $failingScriptPath = Join-Path $tempRoot "exit-23.ps1"
    $passingScriptPath = Join-Path $tempRoot "exit-0.ps1"
    Set-Content -LiteralPath $failingScriptPath -Value "exit 23" -Encoding UTF8
    Set-Content -LiteralPath $passingScriptPath -Value "exit 0" -Encoding UTF8

    $caughtExpectedFailure = $false
    try {
        Invoke-CheckedPowerShellScript `
            -ScriptPath $failingScriptPath `
            -Arguments @() `
            -Description "MutationFailure"
    }
    catch {
        $caughtExpectedFailure = $_.Exception.Message -match 'exit code 23'
    }

    if (-not $caughtExpectedFailure) {
        throw "A non-zero cleanup command did not cause the final cleanup helper to fail."
    }

    Invoke-CheckedPowerShellScript `
        -ScriptPath $passingScriptPath `
        -Arguments @() `
        -Description "MutationSuccess"

    $injectedCleanupErrors = New-Object System.Collections.Generic.List[string]
    $injectedCleanupErrors.Add("UninstallRelease: exit code 23") | Out-Null
    $injectedCleanupErrors.Add("RegisterRepositoryAddIns: exit code 31") | Out-Null
    $injectedCleanupErrors.Add("RestoreSettings: simulated restore failure") | Out-Null
    $injectedDetails = @(
        Get-FullE2EFailureDetails `
            -FatalError ([InvalidOperationException]::new("MainFailure")) `
            -CleanupErrors $injectedCleanupErrors)
    if ($injectedDetails.Count -ne 4) {
        throw "Failure aggregation did not preserve the main error and all cleanup errors."
    }

    $injectedResults = New-Object System.Collections.Generic.List[string]
    Add-FullE2EFailureResult `
        -Results $injectedResults `
        -FailureDetails $injectedDetails
    $injectedLine = $injectedResults -join "`n"
    foreach ($expectedDetail in @(
        "MainFailure",
        "UninstallRelease: exit code 23",
        "RegisterRepositoryAddIns: exit code 31",
        "RestoreSettings: simulated restore failure")) {
        if ($injectedLine -notlike ("*" + $expectedDetail + "*")) {
            throw "FatalError report row omitted: $expectedDetail"
        }
    }

    $probeScriptPath = Join-Path $tempRoot "final-exit-probe.ps1"
    $probeReportPath = Join-Path $tempRoot "final-exit-probe-report.md"
    $probeSource = @(
        (Get-FunctionDefinitionText -Ast $ast -Name "Get-ErrorDetail"),
        (Get-FunctionDefinitionText -Ast $ast -Name "Get-FullE2EFailureDetails"),
        (Get-FunctionDefinitionText -Ast $ast -Name "Add-FullE2EFailureResult"),
        '$results = New-Object System.Collections.Generic.List[string]',
        '$cleanupErrors = New-Object System.Collections.Generic.List[string]',
        '$cleanupErrors.Add("UninstallRelease: exit code 23") | Out-Null',
        '$cleanupErrors.Add("RegisterRepositoryAddIns: exit code 31") | Out-Null',
        '$cleanupErrors.Add("RestoreSettings: simulated restore failure") | Out-Null',
        '$failureDetails = @(Get-FullE2EFailureDetails -FatalError ([InvalidOperationException]::new("MainFailure")) -CleanupErrors $cleanupErrors)',
        'Add-FullE2EFailureResult -Results $results -FailureDetails $failureDetails',
        '$results | Set-Content -LiteralPath $args[0] -Encoding UTF8',
        'if ($failureDetails.Count -gt 0) { exit 1 }',
        'throw "Probe unexpectedly had no failure details."'
    ) -join "`r`n"
    Set-Content -LiteralPath $probeScriptPath -Value $probeSource -Encoding UTF8
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probeScriptPath $probeReportPath
    $probeExitCode = [int]$LASTEXITCODE
    if ($probeExitCode -ne 1) {
        throw "The controlled final branch returned $probeExitCode instead of 1."
    }

    $probeReport = Get-Content -Raw -LiteralPath $probeReportPath
    foreach ($expectedDetail in @(
        "MainFailure",
        "UninstallRelease: exit code 23",
        "RegisterRepositoryAddIns: exit code 31",
        "RestoreSettings: simulated restore failure")) {
        if ($probeReport -notlike ("*" + $expectedDetail + "*")) {
            throw "Controlled final report omitted: $expectedDetail"
        }
    }

    $checkedInvocationCount = ([regex]::Matches($fullE2ESource, 'Invoke-CheckedPowerShellScript')).Count
    if ($checkedInvocationCount -lt 3 -or
        $fullE2ESource -notmatch 'UninstallRelease' -or
        $fullE2ESource -notmatch 'RegisterRepositoryAddIns') {
        throw "Full E2E does not route both uninstall and repository registration through checked cleanup commands."
    }

    if ($fullE2ESource -notmatch '\$cleanupErrors\.Add\("UninstallRelease:' -or
        $fullE2ESource -notmatch '\$cleanupErrors\.Add\("RegisterRepositoryAddIns:' -or
        $fullE2ESource -notmatch '(?s)if \(\$failureDetails\.Count -gt 0\).*?Write-Error.*?exit 1') {
        throw "Cleanup command failures do not propagate to the final full E2E exit code."
    }

    Write-Output "FULL_E2E_CLEANUP_SAFETY_TEST_PASS"
}
finally {
    if ($null -ne $dummyProcess -and -not $dummyProcess.HasExited) {
        Stop-Process -Id $dummyProcess.Id -Force -ErrorAction SilentlyContinue
        $dummyProcess.WaitForExit(5000) | Out-Null
    }

    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}
