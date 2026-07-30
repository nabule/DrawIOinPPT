param()

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$fullE2EPath = Join-Path $PSScriptRoot "full-e2e-test.ps1"
$complexMetadataPath = Join-Path $PSScriptRoot "word-complex-metadata-e2e.ps1"
$powerPointRegisterPath = Join-Path $PSScriptRoot "register-addin.ps1"
$powerPointUnregisterPath = Join-Path $PSScriptRoot "unregister-addin.ps1"
$wordRegisterPath = Join-Path $PSScriptRoot "register-word-addin.ps1"
$wordUnregisterPath = Join-Path $PSScriptRoot "unregister-word-addin.ps1"
$tempRoot = Join-Path $env:TEMP ("DrawioPpt\full-e2e-cleanup-safety-" + [Guid]::NewGuid().ToString("N"))
$registryTestRoot = "Software\Greensoft\DrawioPptTests\full-e2e-cleanup-" + [Guid]::NewGuid().ToString("N")
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
    Import-FunctionDefinition -Ast $ast -Name "ConvertTo-ReportSafeDetail"
    Import-FunctionDefinition -Ast $ast -Name "Get-FullE2EFailureDetails"
    Import-FunctionDefinition -Ast $ast -Name "Add-FullE2EFailureResult"
    Import-FunctionDefinition -Ast $ast -Name "Get-RegistryKeyTreeSnapshot"
    Import-FunctionDefinition -Ast $ast -Name "Restore-RegistryKeyTree"
    Import-FunctionDefinition -Ast $ast -Name "Get-OfficeAddInRegistrationSubKeys"
    Import-FunctionDefinition -Ast $ast -Name "Get-OfficeAddInRegistrationSnapshot"
    Import-FunctionDefinition -Ast $ast -Name "Restore-OfficeAddInRegistration"
    Import-FunctionDefinition -Ast $ast -Name "Assert-NoOfficeProcesses"

    $officeSnapshotCommand = Get-Command Get-OfficeAddInRegistrationSnapshot
    if (-not $officeSnapshotCommand.Parameters.ContainsKey("SubKeys")) {
        throw "Office registration snapshot wrapper must allow isolated subkeys to be injected for safety testing."
    }

    $expectedRegistrationSubKeys = @(
        "Software\Microsoft\Office\PowerPoint\Addins\Greensoft.DrawioPptAddIn",
        "Software\Microsoft\Office\PowerPoint\Addins\DrawioPpt.PowerPointAddIn.Connect",
        "Software\Classes\Greensoft.DrawioPptAddIn",
        "Software\Classes\CLSID\{0B8996D8-D6B9-4D61-8E8C-6F2081BFEA31}",
        "Software\Microsoft\Office\Word\Addins\Greensoft.DrawioWordAddIn",
        "Software\Microsoft\Office\Word\Addins\DrawioPpt.WordAddIn.Connect",
        "Software\Classes\Greensoft.DrawioWordAddIn",
        "Software\Classes\CLSID\{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}"
    )
    $actualRegistrationSubKeys = @(Get-OfficeAddInRegistrationSubKeys)
    $registrationSubKeyDifferences = @(
        Compare-Object `
            -ReferenceObject ($expectedRegistrationSubKeys | Sort-Object) `
            -DifferenceObject ($actualRegistrationSubKeys | Sort-Object))
    if ($actualRegistrationSubKeys.Count -ne $expectedRegistrationSubKeys.Count -or
        $registrationSubKeyDifferences.Count -ne 0) {
        throw "Office add-in registration snapshot does not cover every registry tree mutated by install/uninstall."
    }

    $scriptDerivedRegistrationSubKeys = New-Object System.Collections.Generic.List[string]
    foreach ($scriptSpecification in @(
            [pscustomobject]@{
                HostName = "PowerPoint"
                Paths = @($powerPointRegisterPath, $powerPointUnregisterPath)
            },
            [pscustomobject]@{
                HostName = "Word"
                Paths = @($wordRegisterPath, $wordUnregisterPath)
            })) {
        foreach ($registrationScriptPath in $scriptSpecification.Paths) {
            $registrationScriptSource = Get-Content -Raw -LiteralPath $registrationScriptPath
            $removeTreeCallCount = ([regex]::Matches(
                    $registrationScriptSource,
                    'Remove-RegistryKeyTree\s+-SubKey')).Count
            if ($removeTreeCallCount -ne 4) {
                throw "$registrationScriptPath mutates $removeTreeCallCount registration roots; update the full E2E snapshot boundary."
            }

            $assignmentValues = @{}
            foreach ($variableName in @("progId", "className", "clsid")) {
                $assignmentMatch = [regex]::Match(
                    $registrationScriptSource,
                    '(?m)^\$' + $variableName + '\s*=\s*"([^"]+)"')
                if (-not $assignmentMatch.Success) {
                    throw "$registrationScriptPath does not expose a literal $variableName registration assignment."
                }
                $assignmentValues[$variableName] = $assignmentMatch.Groups[1].Value
            }

            $scriptDerivedRegistrationSubKeys.Add(
                "Software\Microsoft\Office\$($scriptSpecification.HostName)\Addins\$($assignmentValues.progId)") | Out-Null
            $scriptDerivedRegistrationSubKeys.Add(
                "Software\Microsoft\Office\$($scriptSpecification.HostName)\Addins\$($assignmentValues.className)") | Out-Null
            $scriptDerivedRegistrationSubKeys.Add(
                "Software\Classes\$($assignmentValues.progId)") | Out-Null
            $scriptDerivedRegistrationSubKeys.Add(
                "Software\Classes\CLSID\$($assignmentValues.clsid)") | Out-Null
        }
    }

    $scriptRegistrationSubKeyDifferences = @(
        Compare-Object `
            -ReferenceObject ($expectedRegistrationSubKeys | Sort-Object -Unique) `
            -DifferenceObject ($scriptDerivedRegistrationSubKeys | Sort-Object -Unique))
    if ($scriptRegistrationSubKeyDifferences.Count -ne 0) {
        throw "Office add-in registration snapshot targets drifted from the register/unregister scripts."
    }

    $existingRegistrySubKey = "$registryTestRoot\Existing"
    $missingRegistrySubKey = "$registryTestRoot\Missing"
    $existingRegistryKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($existingRegistrySubKey)
    try {
        $existingRegistryKey.SetValue(
            "",
            "original-default",
            [Microsoft.Win32.RegistryValueKind]::String)
        $existingRegistryKey.SetValue(
            "LoadBehavior",
            3,
            [Microsoft.Win32.RegistryValueKind]::DWord)
    }
    finally {
        $existingRegistryKey.Dispose()
    }

    $existingChildKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("$existingRegistrySubKey\Child")
    try {
        $existingChildKey.SetValue(
            "Paths",
            [string[]]@("alpha", "beta"),
            [Microsoft.Win32.RegistryValueKind]::MultiString)
    }
    finally {
        $existingChildKey.Dispose()
    }

    $officeRegistrationSnapshots = @(
        Get-OfficeAddInRegistrationSnapshot `
            -SubKeys @($existingRegistrySubKey, $missingRegistrySubKey))

    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($existingRegistrySubKey, $false)
    $mutatedExistingKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($existingRegistrySubKey)
    try {
        $mutatedExistingKey.SetValue("Unexpected", "mutation")
    }
    finally {
        $mutatedExistingKey.Dispose()
    }
    $mutatedMissingKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($missingRegistrySubKey)
    $mutatedMissingKey.Dispose()

    Restore-OfficeAddInRegistration -Snapshots $officeRegistrationSnapshots

    $restoredRegistryKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($existingRegistrySubKey)
    try {
        if ($null -eq $restoredRegistryKey -or
            $restoredRegistryKey.GetValue("") -ne "original-default" -or
            $restoredRegistryKey.GetValueKind("") -ne [Microsoft.Win32.RegistryValueKind]::String -or
            $restoredRegistryKey.GetValueKind("LoadBehavior") -ne [Microsoft.Win32.RegistryValueKind]::DWord -or
            [int]$restoredRegistryKey.GetValue("LoadBehavior") -ne 3 -or
            $null -ne $restoredRegistryKey.GetValue("Unexpected")) {
            throw "Existing registry state was not restored exactly."
        }
    }
    finally {
        if ($null -ne $restoredRegistryKey) {
            $restoredRegistryKey.Dispose()
        }
    }

    $restoredChildKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("$existingRegistrySubKey\Child")
    try {
        if ($null -eq $restoredChildKey) {
            throw "Nested registry key was not restored."
        }

        $restoredPaths = @($restoredChildKey.GetValue("Paths"))
        if ($restoredChildKey.GetValueKind("Paths") -ne [Microsoft.Win32.RegistryValueKind]::MultiString -or
            $restoredPaths.Count -ne 2 -or
            $restoredPaths[0] -ne "alpha" -or
            $restoredPaths[1] -ne "beta") {
            throw "Nested registry values were not restored exactly."
        }
    }
    finally {
        if ($null -ne $restoredChildKey) {
            $restoredChildKey.Dispose()
        }
    }

    $unexpectedMissingKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($missingRegistrySubKey)
    if ($null -ne $unexpectedMissingKey) {
        $unexpectedMissingKey.Dispose()
        throw "A registry key that was absent before the test was not removed during restoration."
    }

    $continuedRegistrySubKey = "$registryTestRoot\ContinuedAfterFailure"
    $continuedRegistrySnapshot = Get-RegistryKeyTreeSnapshot -SubKey $continuedRegistrySubKey
    $continuedMutationKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($continuedRegistrySubKey)
    $continuedMutationKey.Dispose()
    $invalidRegistrySnapshot = [pscustomobject]@{
        SubKey = "$registryTestRoot\Invalid"
        Exists = $true
        Values = @(
            [pscustomobject]@{
                Name = "InvalidBinary"
                Kind = [Microsoft.Win32.RegistryValueKind]::Binary
                Value = "not-binary-data"
            }
        )
        SubKeys = @()
    }
    $aggregateRestoreFailure = $null
    try {
        Restore-OfficeAddInRegistration `
            -Snapshots @($invalidRegistrySnapshot, $continuedRegistrySnapshot)
    }
    catch {
        $aggregateRestoreFailure = $_
    }

    if ($null -eq $aggregateRestoreFailure -or
        $aggregateRestoreFailure.Exception.Message -notlike ("*" + $invalidRegistrySnapshot.SubKey + "*")) {
        throw "Registration restoration did not aggregate the failing registry root."
    }
    $continuedMutationAfterRestore = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($continuedRegistrySubKey)
    if ($null -ne $continuedMutationAfterRestore) {
        $continuedMutationAfterRestore.Dispose()
        throw "A failing registry root prevented later registration roots from being restored."
    }

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

    $residualProcessFailure = $null
    try {
        Assert-NoOfficeProcesses -ProcessNames @("powershell") -TimeoutSeconds 0
    }
    catch {
        $residualProcessFailure = $_
    }
    if ($null -eq $residualProcessFailure -or
        $residualProcessFailure.Exception.Message -notmatch [regex]::Escape("PID=$($dummyProcess.Id)")) {
        throw "Office residual detection did not report the exact remaining process PID."
    }
    if ($dummyProcess.HasExited) {
        throw "Office residual detection terminated a process instead of reporting it."
    }

    Assert-NoOfficeProcesses `
        -ProcessNames @("DrawioPptProcessNameThatCannotExist") `
        -TimeoutSeconds 0

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
    $injectedCleanupErrors.Add("RestoreOfficeAddInRegistration: simulated restore failure") | Out-Null
    $injectedCleanupErrors.Add("OfficeProcessResidual: WINWORD PID=12345") | Out-Null
    $injectedCleanupErrors.Add("RestoreSettings: simulated restore failure") | Out-Null
    $injectedDetails = @(
        Get-FullE2EFailureDetails `
            -FatalError ([InvalidOperationException]::new("MainFailure")) `
            -CleanupErrors $injectedCleanupErrors)
    if ($injectedDetails.Count -ne 5) {
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
        "RestoreOfficeAddInRegistration: simulated restore failure",
        "OfficeProcessResidual: WINWORD PID=12345",
        "RestoreSettings: simulated restore failure")) {
        if ($injectedLine -notlike ("*" + $expectedDetail + "*")) {
            throw "FatalError report row omitted: $expectedDetail"
        }
    }

    $probeScriptPath = Join-Path $tempRoot "final-exit-probe.ps1"
    $probeReportPath = Join-Path $tempRoot "final-exit-probe-report.md"
    $probeSource = @(
        (Get-FunctionDefinitionText -Ast $ast -Name "Get-ErrorDetail"),
        (Get-FunctionDefinitionText -Ast $ast -Name "ConvertTo-ReportSafeDetail"),
        (Get-FunctionDefinitionText -Ast $ast -Name "Get-FullE2EFailureDetails"),
        (Get-FunctionDefinitionText -Ast $ast -Name "Add-FullE2EFailureResult"),
        '$results = New-Object System.Collections.Generic.List[string]',
        '$cleanupErrors = New-Object System.Collections.Generic.List[string]',
        '$cleanupErrors.Add("UninstallRelease: exit code 23") | Out-Null',
        '$cleanupErrors.Add("RestoreOfficeAddInRegistration: simulated restore failure") | Out-Null',
        '$cleanupErrors.Add("OfficeProcessResidual: WINWORD PID=12345") | Out-Null',
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
        "RestoreOfficeAddInRegistration: simulated restore failure",
        "OfficeProcessResidual: WINWORD PID=12345",
        "RestoreSettings: simulated restore failure")) {
        if ($probeReport -notlike ("*" + $expectedDetail + "*")) {
            throw "Controlled final report omitted: $expectedDetail"
        }
    }

    $checkedInvocationCount = ([regex]::Matches($fullE2ESource, 'Invoke-CheckedPowerShellScript')).Count
    if ($checkedInvocationCount -lt 2 -or
        $fullE2ESource -notmatch 'UninstallRelease') {
        throw "Full E2E does not route release uninstall through the checked cleanup command."
    }

    if ($fullE2ESource -notmatch '\$cleanupErrors\.Add\("UninstallRelease:' -or
        $fullE2ESource -notmatch '\$cleanupErrors\.Add\("RestoreOfficeAddInRegistration:' -or
        $fullE2ESource -notmatch '\$cleanupErrors\.Add\("OfficeProcessResidual:' -or
        $fullE2ESource -notmatch 'Restore-OfficeAddInRegistration\s+-Snapshots\s+\$officeAddInRegistrationSnapshot' -or
        $fullE2ESource -notmatch 'Assert-NoOfficeProcesses\s+-ProcessNames\s+@\("WINWORD",\s*"POWERPNT"\)' -or
        $fullE2ESource -notmatch 'ReleaseComObject\(\$addin\)' -or
        $fullE2ESource -notmatch 'ReleaseComObject\(\$comAddIns\)' -or
        $fullE2ESource -match 'RegisterRepositoryAddIns|Stop-RunningPowerPointSilently' -or
        $fullE2ESource -notmatch '(?s)if \(\$failureDetails\.Count -gt 0\).*?Write-Error.*?exit 1') {
        throw "Registration restoration or Office residual failures do not propagate to the final full E2E exit code."
    }

    $snapshotAcquisition = '$officeAddInRegistrationSnapshot = @(Get-OfficeAddInRegistrationSnapshot)'
    $installInvocation = '$installScript = Join-Path $packageRoot "scripts\\install-release.ps1"'
    $snapshotAcquisitionIndex = $fullE2ESource.IndexOf(
        $snapshotAcquisition,
        [System.StringComparison]::Ordinal)
    $installInvocationIndex = $fullE2ESource.IndexOf(
        $installInvocation,
        [System.StringComparison]::Ordinal)
    if ($snapshotAcquisitionIndex -lt 0 -or
        $installInvocationIndex -lt 0 -or
        $snapshotAcquisitionIndex -ge $installInvocationIndex) {
        throw "Office registration must be captured before the temporary release installation changes it."
    }

    $mainTryStatement = $ast.Find(
        {
            param($node)
            $node -is [System.Management.Automation.Language.TryStatementAst] -and
                $null -ne $node.Finally -and
                $node.Body.Extent.Text.IndexOf(
                    $installInvocation,
                    [System.StringComparison]::Ordinal) -ge 0
        },
        $true)
    if ($null -eq $mainTryStatement) {
        throw "Full E2E main try/finally statement was not found."
    }

    $mainFinallySource = $mainTryStatement.Finally.Extent.Text
    if ($mainFinallySource -notmatch 'Restore-OfficeAddInRegistration\s+-Snapshots\s+\$officeAddInRegistrationSnapshot' -or
        $mainFinallySource -notmatch 'Assert-NoOfficeProcesses\s+-ProcessNames\s+@\("WINWORD",\s*"POWERPNT"\)') {
        throw "Registration restoration and the Office residual gate must run from the main finally block."
    }

    Write-Output "FULL_E2E_CLEANUP_SAFETY_TEST_PASS"
}
finally {
    if ($null -ne $dummyProcess -and -not $dummyProcess.HasExited) {
        Stop-Process -Id $dummyProcess.Id -Force -ErrorAction SilentlyContinue
        $dummyProcess.WaitForExit(5000) | Out-Null
    }

    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($registryTestRoot, $false)

    if (Test-Path -LiteralPath $tempRoot) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force
    }
}
