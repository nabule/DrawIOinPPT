param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "Greensoft\DrawioPpt"),
    [int]$DurationSeconds = 12,
    [double]$MaximumP95Ratio = 1.25,
    [double]$MaximumPerRoundP95Ratio = 1.50,
    [double]$MaximumSelectionP95DeltaMs = 100,
    [double]$MaximumAbsoluteSelectionP95Ms = 2000,
    [double]$MaximumAbsoluteP95Ms = 2000,
    [double]$MaximumAbsoluteResizeP95Ms = 750,
    [int]$MinimumOperationsPerRound = 10,
    [int]$WarmupSeconds = 4,
    [int]$ResizeOperationsPerRound = 12,
    [int]$SelectionSettleMilliseconds = 75,
    [int]$SelectionEventTimeoutMilliseconds = 500,
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

if ($DurationSeconds -lt 10) {
    throw "DurationSeconds must be at least 10 seconds for each picture."
}

foreach ($positiveThreshold in @(
        $MaximumP95Ratio,
        $MaximumPerRoundP95Ratio,
        $MaximumSelectionP95DeltaMs,
        $MaximumAbsoluteSelectionP95Ms,
        $MaximumAbsoluteP95Ms,
        $MaximumAbsoluteResizeP95Ms)) {
    if ($positiveThreshold -le 0) {
        throw "Latency thresholds must be greater than zero."
    }
}

if ($MinimumOperationsPerRound -lt 1) {
    throw "MinimumOperationsPerRound must be at least one."
}

if ($WarmupSeconds -lt 1) {
    throw "WarmupSeconds must be at least one."
}

if ($ResizeOperationsPerRound -lt 4 -or
    ($ResizeOperationsPerRound % 4) -ne 0) {
    throw "ResizeOperationsPerRound must be a positive multiple of four."
}

if ($SelectionSettleMilliseconds -lt 1) {
    throw "SelectionSettleMilliseconds must be at least one."
}

if ($SelectionEventTimeoutMilliseconds -lt
    $SelectionSettleMilliseconds) {
    throw "SelectionEventTimeoutMilliseconds must be at least SelectionSettleMilliseconds."
}

$installRootResolved = (Resolve-Path -LiteralPath $InstallRoot).Path
$binRoot = Join-Path $installRootResolved "bin"
$coreAssemblyPath = Join-Path $binRoot "DrawioPpt.Core.dll"
$wordAddInAssemblyPath = Join-Path $binRoot "DrawioPpt.WordAddIn.dll"

foreach ($requiredPath in @($coreAssemblyPath, $wordAddInAssemblyPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required installed file not found: $requiredPath"
    }
}

if (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count -gt 0) {
    throw "Close Word before running the Word UI-thread stress acceptance."
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $env:TEMP "DrawioPpt\word-ui-thread-stress-v1.0.8.json"
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $outputFullPath
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

[void][Reflection.Assembly]::LoadFrom($coreAssemblyPath)
[void][Reflection.Assembly]::LoadFrom($wordAddInAssemblyPath)

if (-not ("DrawioPptWordWindowProcessResolver" -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class DrawioPptWordWindowProcessResolver
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId);

    public static int GetProcessId(long windowHandle)
    {
        uint processId;
        GetWindowThreadProcessId(
            new IntPtr(windowHandle),
            out processId);
        return (int)processId;
    }
}
'@
}

function Get-WordProcessIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application
    )

    $activeWindow = $null
    $process = $null
    try {
        $activeWindow = $Application.ActiveWindow
        $windowHandle = [long]$activeWindow.Hwnd
        $processId = [DrawioPptWordWindowProcessResolver]::GetProcessId(
            $windowHandle)
        if ($processId -le 0) {
            throw "Unable to resolve the Word process from ActiveWindow.Hwnd."
        }

        $process = Get-Process -Id $processId -ErrorAction Stop
        return [pscustomobject]@{
            ProcessId = $process.Id
            StartTimeUtcTicks =
                $process.StartTime.ToUniversalTime().Ticks
        }
    }
    finally {
        if ($null -ne $process) {
            $process.Dispose()
        }

        Release-ComObject -Value $activeWindow
    }
}

function Get-SoleWordProcessIdentity {
    $processes = @(Get-Process WINWORD -ErrorAction SilentlyContinue)
    if ($processes.Count -ne 1) {
        foreach ($process in $processes) {
            $process.Dispose()
        }

        throw "Expected exactly one Word process after creating the isolated COM application, but found $($processes.Count)."
    }

    $process = $processes[0]
    try {
        return [pscustomobject]@{
            ProcessId = $process.Id
            StartTimeUtcTicks =
                $process.StartTime.ToUniversalTime().Ticks
        }
    }
    finally {
        $process.Dispose()
    }
}

function Assert-WordProcessIdentityMatch {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Expected,
        [Parameter(Mandatory = $true)]
        [object]$Actual
    )

    if ($Expected.ProcessId -ne $Actual.ProcessId -or
        $Expected.StartTimeUtcTicks -ne
            $Actual.StartTimeUtcTicks) {
        throw "The Word ActiveWindow process identity does not match the process created for this isolated COM application."
    }
}

function Release-ComObject {
    param(
        [object]$Value
    )

    if ($null -eq $Value -or
        -not [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        return
    }

    try {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
            $Value)
    }
    catch {
    }
}

function Get-WordProcessForIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Identity
    )

    $process = Get-Process `
        -Id $Identity.ProcessId `
        -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        return $null
    }

    try {
        $actualStartTicks =
            $process.StartTime.ToUniversalTime().Ticks
        if ($actualStartTicks -ne $Identity.StartTimeUtcTicks) {
            $process.Dispose()
            return $null
        }

        return $process
    }
    catch {
        $process.Dispose()
        return $null
    }
}

function New-ComplexSvg {
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="800" viewBox="0 0 1200 800">')
    [void]$builder.Append('<rect width="1200" height="800" fill="#f7fbff"/>')
    [void]$builder.Append('<text x="40" y="55" font-family="Segoe UI" font-size="30" fill="#17365d">DrawioPpt v1.0.8 Complex Vector</text>')

    for ($index = 0; $index -lt 120; $index++) {
        $column = $index % 30
        $row = [Math]::Floor($index / 30)
        $x = 22 + ($column * 38)
        $y = 82 + ($row * 34)
        $fill = if (($index % 3) -eq 0) {
            "#5b9bd5"
        }
        elseif (($index % 3) -eq 1) {
            "#70ad47"
        }
        else {
            "#ed7d31"
        }

        [void]$builder.Append(
            ('<g><rect x="{0}" y="{1}" width="30" height="22" rx="4" fill="{2}" stroke="#fff"/>' -f $x, $y, $fill))
        [void]$builder.Append(
            ('<text x="{0}" y="{1}" font-family="Segoe UI" font-size="8" fill="#fff">{2}</text></g>' -f ($x + 4), ($y + 15), $index))
    }

    [void]$builder.Append("</svg>")
    return $builder.ToString()
}

function New-ComplexDrawioXml {
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('<mxfile host="DrawioWord"><diagram id="ui-thread-stress-v108" name="Managed Complex v1.0.8"><mxGraphModel><root><mxCell id="0"/><mxCell id="1" parent="0"/>')

    for ($index = 0; $index -lt 1200; $index++) {
        $x = ($index % 40) * 30
        $y = [Math]::Floor($index / 40) * 25
        [void]$builder.Append(
            ('<mxCell id="n{0}" value="Node {0}" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#dae8fc;strokeColor=#6c8ebf;" vertex="1" parent="1"><mxGeometry x="{1}" y="{2}" width="28" height="20" as="geometry"/></mxCell>' -f $index, $x, $y))
    }

    [void]$builder.Append("</root></mxGraphModel></diagram></mxfile>")
    return $builder.ToString()
}

function Get-ShapeByName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $shapes = $null
    try {
        $shapes = $Document.Shapes
        for ($index = 1; $index -le $shapes.Count; $index++) {
            $candidate = $shapes.Item($index)
            if ([string]::Equals(
                    [string]$candidate.Name,
                    $Name,
                    [System.StringComparison]::Ordinal)) {
                return $candidate
            }

            Release-ComObject -Value $candidate
        }
    }
    finally {
        Release-ComObject -Value $shapes
    }

    return $null
}

function Open-WordDocument {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $documents = $null
    try {
        $documents = $Application.Documents
        return $documents.Open($Path)
    }
    finally {
        Release-ComObject -Value $documents
    }
}

function Get-WordComAddIn {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application,
        [Parameter(Mandatory = $true)]
        [string]$ProgId
    )

    $comAddIns = $null
    try {
        $comAddIns = $Application.COMAddIns
        return $comAddIns.Item($ProgId)
    }
    finally {
        Release-ComObject -Value $comAddIns
    }
}

function Get-Percentile {
    param(
        [Parameter(Mandatory = $true)]
        [double[]]$Values,
        [Parameter(Mandatory = $true)]
        [double]$Percentile
    )

    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 0) {
        return 0
    }

    $index = [Math]::Min(
        $sorted.Count - 1,
        [Math]::Floor($sorted.Count * $Percentile))
    return [double]$sorted[$index]
}

function Wait-ShapeSelectionEvent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SelectionCounter,
        [Parameter(Mandatory = $true)]
        [string]$ShapeName,
        [Parameter(Mandatory = $true)]
        [int]$CountBefore,
        [Parameter(Mandatory = $true)]
        [int]$MinimumWaitMilliseconds,
        [Parameter(Mandatory = $true)]
        [int]$TimeoutMilliseconds
    )

    Start-Sleep -Milliseconds $MinimumWaitMilliseconds
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($SelectionCounter.GetShapeSelectionCount(
            $ShapeName) -le $CountBefore -and
        $watch.Elapsed.TotalMilliseconds -lt
            $TimeoutMilliseconds) {
        Start-Sleep -Milliseconds 10
    }

    return $SelectionCounter.GetShapeSelectionCount(
        $ShapeName) -gt $CountBefore
}

function Measure-ShapeMotion {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,
        [Parameter(Mandatory = $true)]
        [object]$Shape,
        [Parameter(Mandatory = $true)]
        [object]$OtherShape,
        [Parameter(Mandatory = $true)]
        [string]$Label,
        [Parameter(Mandatory = $true)]
        [int]$Seconds,
        [Parameter(Mandatory = $true)]
        [int]$WarmupSeconds,
        [Parameter(Mandatory = $true)]
        [int]$ResizeOperations,
        [Parameter(Mandatory = $true)]
        [int]$SelectionSettleMilliseconds,
        [Parameter(Mandatory = $true)]
        [int]$SelectionEventTimeoutMilliseconds,
        [Parameter(Mandatory = $true)]
        [object]$SelectionCounter,
        [Parameter(Mandatory = $true)]
        [object]$WordProcessIdentity
    )

    $neutralRange = $null
    try {
        $OtherShape.Visible = 0
        $Shape.Visible = -1
        $Shape.Left = 45
        $Shape.Top = 70
        $Shape.Width = 330
        $Shape.Height = 220
        $Shape.LockAspectRatio = 0
        $Shape.Select()
        Start-Sleep -Milliseconds 250

        $neutralRange = $Document.Range(0, 0)
        $warmupWatch = [Diagnostics.Stopwatch]::StartNew()
        $warmupIteration = 0
        while ($warmupWatch.Elapsed.TotalSeconds -lt
            $WarmupSeconds) {
            $warmupIteration++
            $neutralRange.Select()
            $Shape.Select()
            if (($warmupIteration % 2) -eq 0) {
                $Shape.Left = 43
                $Shape.Top = 69
            }
            else {
                $Shape.Left = 47
                $Shape.Top = 71
            }

            Start-Sleep -Milliseconds 20
        }

        $warmupWatch.Stop()
        $Shape.Left = 45
        $Shape.Top = 70
        Start-Sleep -Milliseconds 1000
        $shapeName = [string]$Shape.Name
        $SelectionCounter.ResetShapeSelectionCount(
            $shapeName)
        $shapeSelectionEventsBefore =
            $SelectionCounter.GetShapeSelectionCount(
                $shapeName)
        $baseLeft = [single]$Shape.Left
        $baseTop = [single]$Shape.Top
        $baseWidth = [single]$Shape.Width
        $baseHeight = [single]$Shape.Height
        $samples = New-Object System.Collections.Generic.List[double]
        $selectionSamples =
            New-Object System.Collections.Generic.List[double]
        $selectionFailures = 0
        $confirmedShapeSelectionEvents = 0
        $missingShapeSelectionEvents = 0
        $setFailures = 0
        $transientRetries = 0
        $unresponsiveSamples = 0
        $selectionOperations = 0
        $iteration = 0
        $totalWatch = [Diagnostics.Stopwatch]::StartNew()

        while ($totalWatch.Elapsed.TotalSeconds -lt $Seconds) {
            $iteration++
            $neutralRange.Select()
            Start-Sleep `
                -Milliseconds $SelectionSettleMilliseconds
            $shapeEventCountBeforeSelection =
                $SelectionCounter.GetShapeSelectionCount(
                    $shapeName)
            $selectionWatch =
                [Diagnostics.Stopwatch]::StartNew()
            $selectionSucceeded = $false
            try {
                $Shape.Select()
                $selectionOperations++
                $selectionSucceeded = $true
            }
            catch {
                $selectionFailures++
            }
            finally {
                $selectionWatch.Stop()
                $selectionSamples.Add(
                    $selectionWatch.Elapsed.TotalMilliseconds)
            }

            if ($selectionSucceeded) {
                if (Wait-ShapeSelectionEvent `
                    -SelectionCounter $SelectionCounter `
                    -ShapeName $shapeName `
                    -CountBefore $shapeEventCountBeforeSelection `
                    -MinimumWaitMilliseconds $SelectionSettleMilliseconds `
                    -TimeoutMilliseconds $SelectionEventTimeoutMilliseconds) {
                    $confirmedShapeSelectionEvents++
                }
                else {
                    $missingShapeSelectionEvents++
                }
            }

            if (-not $selectionSucceeded) {
                Start-Sleep -Milliseconds 20
                continue
            }

            $stepWatch =
                [Diagnostics.Stopwatch]::StartNew()
            $setSucceeded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    $phase = $iteration % 4
                    if ($phase -eq 0) {
                        $Shape.Left = $baseLeft - 18
                        $Shape.Top = $baseTop - 8
                    }
                    elseif ($phase -eq 1) {
                        $Shape.Left = $baseLeft + 18
                        $Shape.Top = $baseTop - 8
                    }
                    elseif ($phase -eq 2) {
                        $Shape.Left = $baseLeft + 18
                        $Shape.Top = $baseTop + 8
                    }
                    else {
                        $Shape.Left = $baseLeft - 18
                        $Shape.Top = $baseTop + 8
                    }

                    $setSucceeded = $true
                    break
                }
                catch {
                    if ($attempt -lt 3) {
                        $transientRetries++
                        Start-Sleep -Milliseconds 250
                    }
                }
            }

            if (-not $setSucceeded) {
                $setFailures++
            }

            $stepWatch.Stop()
            $samples.Add($stepWatch.Elapsed.TotalMilliseconds)

            $wordProcess = Get-WordProcessForIdentity `
                -Identity $WordProcessIdentity
            if ($null -eq $wordProcess) {
                throw "The test-owned Word process exited or its identity changed."
            }

            try {
                $wordProcess.Refresh()
                if (-not $wordProcess.Responding) {
                    $unresponsiveSamples++
                }
            }
            finally {
                $wordProcess.Dispose()
            }

            Start-Sleep -Milliseconds 20
        }

        $totalWatch.Stop()
        Start-Sleep -Milliseconds 500

        $resizeSamples = New-Object System.Collections.Generic.List[double]
        $resizeTransientRetries = 0
        $resizeFailures = 0
        for ($resizeIndex = 0;
            $resizeIndex -lt $ResizeOperations;
            $resizeIndex++) {
            $neutralRange.Select()
            Start-Sleep `
                -Milliseconds $SelectionSettleMilliseconds
            $shapeEventCountBeforeResizeSelection =
                $SelectionCounter.GetShapeSelectionCount(
                    $shapeName)
            $resizeSelectionWatch =
                [Diagnostics.Stopwatch]::StartNew()
            $resizeSelectionSucceeded = $false
            try {
                $Shape.Select()
                $selectionOperations++
                $resizeSelectionSucceeded = $true
            }
            catch {
                $selectionFailures++
            }
            finally {
                $resizeSelectionWatch.Stop()
                $selectionSamples.Add(
                    $resizeSelectionWatch.Elapsed.TotalMilliseconds)
            }

            if ($resizeSelectionSucceeded) {
                if (Wait-ShapeSelectionEvent `
                    -SelectionCounter $SelectionCounter `
                    -ShapeName $shapeName `
                    -CountBefore $shapeEventCountBeforeResizeSelection `
                    -MinimumWaitMilliseconds $SelectionSettleMilliseconds `
                    -TimeoutMilliseconds $SelectionEventTimeoutMilliseconds) {
                    $confirmedShapeSelectionEvents++
                }
                else {
                    $missingShapeSelectionEvents++
                }
            }

            if (-not $resizeSelectionSucceeded) {
                $resizeFailures++
                continue
            }

            $resizeWatch = [Diagnostics.Stopwatch]::StartNew()
            $resizeSucceeded = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                try {
                    $resizePhase = $resizeIndex % 4
                    if ($resizePhase -eq 0) {
                        $Shape.Width = [single]($baseWidth + 6)
                    }
                    elseif ($resizePhase -eq 1) {
                        $Shape.Width = [single]$baseWidth
                    }
                    elseif ($resizePhase -eq 2) {
                        $Shape.Height = [single]($baseHeight + 4)
                    }
                    else {
                        $Shape.Height = [single]$baseHeight
                    }

                    $resizeSucceeded = $true
                    break
                }
                catch {
                    if ($attempt -lt 3) {
                        $resizeTransientRetries++
                        Start-Sleep -Milliseconds 250
                    }
                }
            }

            if (-not $resizeSucceeded) {
                $resizeFailures++
            }

            $resizeWatch.Stop()
            $resizeSamples.Add($resizeWatch.Elapsed.TotalMilliseconds)
            Start-Sleep -Milliseconds 250
        }

        Start-Sleep -Milliseconds 500
        $shapeSelectionEventsAfter =
            $SelectionCounter.GetShapeSelectionCount(
                $shapeName)
        $classifiedShapeSelectionEventDelta =
            [Math]::Max(
                0,
                $shapeSelectionEventsAfter -
                    $shapeSelectionEventsBefore)

        $sampleArray = [double[]]$samples.ToArray()
        $selectionSampleArray =
            [double[]]$selectionSamples.ToArray()
        $resizeSampleArray = [double[]]$resizeSamples.ToArray()
        $average = ($sampleArray | Measure-Object -Average).Average
        $maximum = ($sampleArray | Measure-Object -Maximum).Maximum
        $resizeAverage = ($resizeSampleArray | Measure-Object -Average).Average
        $resizeMaximum = ($resizeSampleArray | Measure-Object -Maximum).Maximum
        $selectionAverage =
            ($selectionSampleArray | Measure-Object -Average).Average
        $selectionMaximum =
            ($selectionSampleArray | Measure-Object -Maximum).Maximum

        return [pscustomobject]@{
            Label = $Label
            DurationSeconds = [Math]::Round(
                $totalWatch.Elapsed.TotalSeconds,
                3)
            Operations = $sampleArray.Count
            SelectionOperations = $selectionOperations
            SelectionAverageMs = [Math]::Round(
                $selectionAverage,
                3)
            SelectionP95Ms = [Math]::Round(
                (Get-Percentile `
                    -Values $selectionSampleArray `
                    -Percentile 0.95),
                3)
            SelectionMaxMs = [Math]::Round(
                $selectionMaximum,
                3)
            SelectionFailures = $selectionFailures
            ConfirmedShapeSelectionEvents =
                $confirmedShapeSelectionEvents
            MissingShapeSelectionEvents =
                $missingShapeSelectionEvents
            ClassifiedShapeSelectionEventDelta =
                $classifiedShapeSelectionEventDelta
            AverageMs = [Math]::Round($average, 3)
            P95Ms = [Math]::Round(
                (Get-Percentile -Values $sampleArray -Percentile 0.95),
                3)
            P99Ms = [Math]::Round(
                (Get-Percentile -Values $sampleArray -Percentile 0.99),
                3)
            MaxMs = [Math]::Round($maximum, 3)
            Over100Ms = @($sampleArray | Where-Object { $_ -gt 100 }).Count
            Over250Ms = @($sampleArray | Where-Object { $_ -gt 250 }).Count
            TransientRetries = $transientRetries
            SetFailures = $setFailures
            ResizeOperations = $resizeSampleArray.Count
            ResizeAverageMs = [Math]::Round($resizeAverage, 3)
            ResizeP95Ms = [Math]::Round(
                (Get-Percentile -Values $resizeSampleArray -Percentile 0.95),
                3)
            ResizeMaxMs = [Math]::Round($resizeMaximum, 3)
            ResizeTransientRetries = $resizeTransientRetries
            ResizeSetFailures = $resizeFailures
            UnresponsiveSamples = $unresponsiveSamples
            StartLeft = $baseLeft
            StartTop = $baseTop
            SelectionSamplesMs = $selectionSampleArray
            MotionSamplesMs = $sampleArray
            ResizeSamplesMs = $resizeSampleArray
        }
    }
    finally {
        Release-ComObject -Value $neutralRange
    }
}

function Merge-ShapeMotionResults {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Results,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $sampleArray = [double[]]@(
        $Results | ForEach-Object { $_.MotionSamplesMs })
    $selectionSampleArray = [double[]]@(
        $Results | ForEach-Object { $_.SelectionSamplesMs })
    $resizeSampleArray = [double[]]@(
        $Results | ForEach-Object { $_.ResizeSamplesMs })

    return [pscustomobject]@{
        Label = $Label
        Rounds = $Results.Count
        DurationSeconds = [Math]::Round(
            ($Results | Measure-Object -Property DurationSeconds -Sum).Sum,
            3)
        Operations =
            ($Results | Measure-Object -Property Operations -Sum).Sum
        SelectionOperations =
            ($Results | Measure-Object -Property SelectionOperations -Sum).Sum
        SelectionAverageMs = [Math]::Round(
            ($selectionSampleArray |
                Measure-Object -Average).Average,
            3)
        SelectionP95Ms = [Math]::Round(
            (Get-Percentile `
                -Values $selectionSampleArray `
                -Percentile 0.95),
            3)
        SelectionMaxMs = [Math]::Round(
            ($selectionSampleArray |
                Measure-Object -Maximum).Maximum,
            3)
        SelectionFailures =
            ($Results |
                Measure-Object -Property SelectionFailures -Sum).Sum
        ConfirmedShapeSelectionEvents =
            ($Results |
                Measure-Object `
                    -Property ConfirmedShapeSelectionEvents `
                    -Sum).Sum
        MissingShapeSelectionEvents =
            ($Results |
                Measure-Object `
                    -Property MissingShapeSelectionEvents `
                    -Sum).Sum
        ClassifiedShapeSelectionEventDelta =
            ($Results |
                Measure-Object `
                    -Property ClassifiedShapeSelectionEventDelta `
                    -Sum).Sum
        AverageMs = [Math]::Round(
            ($sampleArray | Measure-Object -Average).Average,
            3)
        P95Ms = [Math]::Round(
            (Get-Percentile -Values $sampleArray -Percentile 0.95),
            3)
        P99Ms = [Math]::Round(
            (Get-Percentile -Values $sampleArray -Percentile 0.99),
            3)
        MaxMs = [Math]::Round(
            ($sampleArray | Measure-Object -Maximum).Maximum,
            3)
        Over100Ms =
            ($Results | Measure-Object -Property Over100Ms -Sum).Sum
        Over250Ms =
            ($Results | Measure-Object -Property Over250Ms -Sum).Sum
        TransientRetries =
            ($Results | Measure-Object -Property TransientRetries -Sum).Sum
        SetFailures =
            ($Results | Measure-Object -Property SetFailures -Sum).Sum
        ResizeOperations =
            ($Results | Measure-Object -Property ResizeOperations -Sum).Sum
        ResizeAverageMs = [Math]::Round(
            ($resizeSampleArray | Measure-Object -Average).Average,
            3)
        ResizeP95Ms = [Math]::Round(
            (Get-Percentile -Values $resizeSampleArray -Percentile 0.95),
            3)
        ResizeMaxMs = [Math]::Round(
            ($resizeSampleArray | Measure-Object -Maximum).Maximum,
            3)
        ResizeTransientRetries =
            ($Results |
                Measure-Object -Property ResizeTransientRetries -Sum).Sum
        ResizeSetFailures =
            ($Results |
                Measure-Object -Property ResizeSetFailures -Sum).Sum
        UnresponsiveSamples =
            ($Results |
                Measure-Object -Property UnresponsiveSamples -Sum).Sum
        StartLeft = [single]$Results[0].StartLeft
        StartTop = [single]$Results[0].StartTop
    }
}

function Wait-WordProcessExit {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Identity,
        [int]$TimeoutSeconds = 15
    )

    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $process = Get-WordProcessForIdentity -Identity $Identity
        if ($null -eq $process) {
            return $true
        }

        $process.Dispose()
        Start-Sleep -Milliseconds 250
    }

    return $false
}

function Stop-OwnedWordProcess {
    param(
        [object]$Identity
    )

    if ($null -eq $Identity) {
        return
    }

    $process = Get-WordProcessForIdentity -Identity $Identity
    if ($null -eq $process) {
        return
    }

    try {
        Stop-Process -Id $Identity.ProcessId -Force -ErrorAction Stop
        if (-not $process.WaitForExit(5000)) {
            throw "The test-owned Word process did not exit after forced cleanup."
        }
    }
    finally {
        $process.Dispose()
    }

    $remainingProcess =
        Get-WordProcessForIdentity -Identity $Identity
    if ($null -ne $remainingProcess) {
        $remainingProcess.Dispose()
        throw "The test-owned Word process identity still exists after forced cleanup."
    }
}

function New-TypedStressDocument {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,
        [Parameter(Mandatory = $true)]
        [string]$DocumentPath,
        [Parameter(Mandatory = $true)]
        [string]$DrawioXml,
        [Parameter(Mandatory = $true)]
        [string]$CoreAssemblyPath,
        [Parameter(Mandatory = $true)]
        [string]$WordAddInAssemblyPath
    )

    $wordInteropPath = (
        Get-ChildItem `
            -LiteralPath "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Word" `
            -Recurse `
            -Filter "Microsoft.Office.Interop.Word.dll" `
            -ErrorAction Stop |
            Select-Object -First 1).FullName
    $officeInteropPath = (
        Get-ChildItem `
            -LiteralPath "C:\Windows\assembly\GAC_MSIL\Office" `
            -Recurse `
            -Filter "OFFICE.DLL" `
            -ErrorAction Stop |
            Select-Object -First 1).FullName

    if ([string]::IsNullOrWhiteSpace($wordInteropPath) -or
        [string]::IsNullOrWhiteSpace($officeInteropPath)) {
        throw "Microsoft Office interop assemblies were not found in the GAC."
    }

    if (-not ("DrawioPptWordUiStressDocumentBuilder" -as [type])) {
        $source = @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;
using DrawioPpt.WordAddIn.Services;
using DrawioPpt.WordAddIn.Word;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

public sealed class DrawioPptWordSelectionChangeCounter :
    IDisposable
{
    private Word.Application application;

    public int Count { get; private set; }
    public int ManagedShapeSelectionCount { get; private set; }
    public int PlainShapeSelectionCount { get; private set; }

    private static void ReleaseComObject(object value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            try
            {
                Marshal.FinalReleaseComObject(value);
            }
            catch
            {
            }
        }
    }

    public int GetShapeSelectionCount(string shapeName)
    {
        if (string.Equals(
            shapeName,
            "Managed Complex v1.0.8",
            StringComparison.Ordinal))
        {
            return ManagedShapeSelectionCount;
        }

        if (string.Equals(
            shapeName,
            "Plain Complex Same Visual",
            StringComparison.Ordinal))
        {
            return PlainShapeSelectionCount;
        }

        throw new ArgumentException(
            "Unexpected stress-test shape name.",
            "shapeName");
    }

    public void ResetShapeSelectionCount(string shapeName)
    {
        if (string.Equals(
            shapeName,
            "Managed Complex v1.0.8",
            StringComparison.Ordinal))
        {
            ManagedShapeSelectionCount = 0;
            return;
        }

        if (string.Equals(
            shapeName,
            "Plain Complex Same Visual",
            StringComparison.Ordinal))
        {
            PlainShapeSelectionCount = 0;
            return;
        }

        throw new ArgumentException(
            "Unexpected stress-test shape name.",
            "shapeName");
    }

    public void Attach(Word.Application target)
    {
        if (target == null)
        {
            throw new ArgumentNullException("target");
        }

        Detach();
        application = target;
        application.WindowSelectionChange +=
            OnWindowSelectionChange;
    }

    private void OnWindowSelectionChange(
        Word.Selection selection)
    {
        Count++;
        Word.ShapeRange shapeRange = null;
        Word.Shape shape = null;
        try
        {
            shapeRange = selection.ShapeRange;
            shape = shapeRange[1];
            string shapeName = shape.Name;
            if (string.Equals(
                shapeName,
                "Managed Complex v1.0.8",
                StringComparison.Ordinal))
            {
                ManagedShapeSelectionCount++;
            }
            else if (string.Equals(
                shapeName,
                "Plain Complex Same Visual",
                StringComparison.Ordinal))
            {
                PlainShapeSelectionCount++;
            }
        }
        catch
        {
        }
        finally
        {
            ReleaseComObject(shape);
            ReleaseComObject(shapeRange);
        }
    }

    public void Detach()
    {
        if (application == null)
        {
            return;
        }

        try
        {
            application.WindowSelectionChange -=
                OnWindowSelectionChange;
        }
        finally
        {
            application = null;
        }
    }

    public void Dispose()
    {
        Detach();
    }
}

public static class DrawioPptWordUiStressDocumentBuilder
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId);

    private static void ReleaseComObject(object value)
    {
        if (value != null && Marshal.IsComObject(value))
        {
            try
            {
                Marshal.FinalReleaseComObject(value);
            }
            catch
            {
            }
        }
    }

    private static bool IsOwnedProcessRunning(
        uint processId,
        long processStartTimeUtcTicks)
    {
        if (processId == 0 || processStartTimeUtcTicks == 0)
        {
            return false;
        }

        try
        {
            using (Process process = Process.GetProcessById((int)processId))
            {
                return process.StartTime.ToUniversalTime().Ticks ==
                    processStartTimeUtcTicks;
            }
        }
        catch
        {
            return false;
        }
    }

    private static void CaptureSoleWordProcess(
        out uint processId,
        out long processStartTimeUtcTicks)
    {
        Process[] processes =
            Process.GetProcessesByName("WINWORD");
        try
        {
            if (processes.Length != 1)
            {
                throw new InvalidOperationException(
                    "Expected exactly one isolated Word process after creating the document-builder COM application, but found " +
                    processes.Length.ToString() + ".");
            }

            processId = (uint)processes[0].Id;
            processStartTimeUtcTicks =
                processes[0].StartTime.ToUniversalTime().Ticks;
        }
        finally
        {
            foreach (Process process in processes)
            {
                process.Dispose();
            }
        }
    }

    private static void StopOwnedProcessIfNeeded(
        uint processId,
        long processStartTimeUtcTicks)
    {
        Stopwatch watch = Stopwatch.StartNew();
        while (watch.Elapsed.TotalSeconds < 15 &&
            IsOwnedProcessRunning(processId, processStartTimeUtcTicks))
        {
            Thread.Sleep(250);
        }

        if (!IsOwnedProcessRunning(processId, processStartTimeUtcTicks))
        {
            return;
        }

        using (Process process = Process.GetProcessById((int)processId))
        {
            if (process.StartTime.ToUniversalTime().Ticks !=
                processStartTimeUtcTicks)
            {
                return;
            }

            process.Kill();
            if (!process.WaitForExit(5000) ||
                IsOwnedProcessRunning(
                    processId,
                    processStartTimeUtcTicks))
            {
                throw new InvalidOperationException(
                    "The document-builder Word process remained after forced cleanup.");
            }
        }
    }

    private static Word.Shape AddPicture(
        Word.Document document,
        string imagePath,
        float top,
        string name)
    {
        object linkToFile = false;
        object saveWithDocument = true;
        Word.Range anchor = null;
        Word.InlineShapes inlineShapes = null;
        Word.InlineShape inline = null;
        Word.WrapFormat wrapFormat = null;
        try
        {
            anchor = document.Range(0, 0);
            inlineShapes = document.InlineShapes;
            inline = inlineShapes.AddPicture(
                imagePath,
                ref linkToFile,
                ref saveWithDocument,
                anchor);
            Word.Shape shape = inline.ConvertToShape();
            shape.Name = name;
            shape.RelativeHorizontalPosition =
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            shape.RelativeVerticalPosition =
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            shape.Left = 45f;
            shape.Top = top;
            shape.Width = 330f;
            shape.Height = 220f;
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            wrapFormat = shape.WrapFormat;
            wrapFormat.Type = Word.WdWrapType.wdWrapSquare;
            return shape;
        }
        finally
        {
            ReleaseComObject(wrapFormat);
            ReleaseComObject(inline);
            ReleaseComObject(inlineShapes);
            ReleaseComObject(anchor);
        }
    }

    public static string Build(
        string imagePath,
        string documentPath,
        string drawioXml)
    {
        Word.Application application = null;
        Word.Documents documents = null;
        Word.Document document = null;
        Word.PageSetup pageSetup = null;
        Word.Shape managed = null;
        Word.Shape plain = null;
        uint processId = 0;
        long processStartTimeUtcTicks = 0;
        try
        {
            application = new Word.Application();
            CaptureSoleWordProcess(
                out processId,
                out processStartTimeUtcTicks);
            application.Visible = false;
            application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            documents = application.Documents;
            document = documents.Add();
            ReleaseComObject(documents);
            documents = null;
            object activeWindowObject = null;
            try
            {
                dynamic activeWindow = application.ActiveWindow;
                activeWindowObject = activeWindow;
                long windowHandle = (long)activeWindow.Hwnd;
                uint activeWindowProcessId;
                GetWindowThreadProcessId(
                    new IntPtr(windowHandle),
                    out activeWindowProcessId);
                if (activeWindowProcessId != processId)
                {
                    throw new InvalidOperationException(
                        "The document-builder ActiveWindow process does not match the isolated Word process.");
                }
            }
            finally
            {
                ReleaseComObject(activeWindowObject);
            }

            if (processId == 0)
            {
                throw new InvalidOperationException(
                    "Unable to resolve the document-builder Word process.");
            }

            pageSetup = document.PageSetup;
            pageSetup.TopMargin = 30f;
            pageSetup.BottomMargin = 30f;
            pageSetup.LeftMargin = 30f;
            pageSetup.RightMargin = 30f;
            ReleaseComObject(pageSetup);
            pageSetup = null;

            managed = AddPicture(
                document,
                imagePath,
                70f,
                "Managed Complex v1.0.8");
            plain = AddPicture(
                document,
                imagePath,
                390f,
                "Plain Complex Same Visual");

            DiagramEnvelope envelope = new DiagramEnvelope();
            envelope.DiagramId = Guid.NewGuid().ToString("N");
            envelope.DiagramName = "Managed Complex v1.0.8";
            envelope.EditorMode = EditorMode.Url;
            envelope.EditorTarget = "https://app.diagrams.net";
            envelope.SidecarPath = string.Empty;
            envelope.UpdatedUtc = DateTime.UtcNow;
            envelope.DrawioXml = drawioXml;

            DiagramEnvelopeSerializer serializer =
                new DiagramEnvelopeSerializer();
            DocumentDiagramStore store =
                new DocumentDiagramStore(serializer);
            string partId = store.Upsert(document, envelope);
            if (string.IsNullOrWhiteSpace(partId))
            {
                throw new InvalidOperationException(
                    "Document.CustomXMLParts upsert failed.");
            }

            WordPictureMetadataService metadata =
                new WordPictureMetadataService(
                    serializer,
                    new PresentationSidecarPathBuilder());
            metadata.Save(
                WordPictureReference.FromShape(managed),
                envelope);
            plain.Title = "Plain Complex Same Visual";
            plain.AlternativeText = string.Empty;

            string alternativeText = managed.AlternativeText ?? string.Empty;
            DiagramEnvelope stored;
            if (alternativeText.Length > 2048 ||
                !store.TryRead(document, envelope.DiagramId, out stored) ||
                stored == null ||
                stored.DrawioXml != drawioXml)
            {
                throw new InvalidOperationException(
                    "Initial metadata verification failed.");
            }

            document.SaveAs2(documentPath);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            ReleaseComObject(document);
            document = null;
            application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            ReleaseComObject(application);
            application = null;

            return
                processId.ToString() + "|" +
                processStartTimeUtcTicks.ToString() + "|" +
                partId + "|" +
                alternativeText.Length.ToString() + "|" +
                envelope.DiagramId;
        }
        finally
        {
            ReleaseComObject(pageSetup);
            ReleaseComObject(documents);
            ReleaseComObject(plain);
            ReleaseComObject(managed);
            if (document != null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }

                ReleaseComObject(document);
            }

            if (application != null)
            {
                try
                {
                    application.Quit(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }

                ReleaseComObject(application);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            StopOwnedProcessIfNeeded(
                processId,
                processStartTimeUtcTicks);
        }
    }
}
'@

        Add-Type `
            -TypeDefinition $source `
            -ReferencedAssemblies @(
                $CoreAssemblyPath,
                $WordAddInAssemblyPath,
                $wordInteropPath,
                $officeInteropPath,
                "System.dll",
                "System.Core.dll",
                "Microsoft.CSharp.dll",
                "System.Xml.dll",
                "System.Xml.Linq.dll") `
            -Language CSharp `
            -IgnoreWarnings
    }

    $buildText = [DrawioPptWordUiStressDocumentBuilder]::Build(
        $ImagePath,
        $DocumentPath,
        $DrawioXml)
    $parts = $buildText -split "\|"
    if ($parts.Count -ne 5) {
        throw "Unexpected typed document builder result: $buildText"
    }

    return [pscustomobject]@{
        ProcessId = [int]$parts[0]
        ProcessStartTimeUtcTicks = [long]$parts[1]
        CustomXmlPartId = $parts[2]
        AlternativeTextChars = [int]$parts[3]
        DiagramId = $parts[4]
    }
}

$testRoot = Join-Path (
    Join-Path $env:TEMP "DrawioPpt") (
    "word-ui-thread-stress-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
$svgPath = Join-Path $testRoot "complex-vector.svg"
$documentPath = Join-Path $testRoot "word-ui-thread-stress.docx"
$complexSvg = New-ComplexSvg
$drawioXml = New-ComplexDrawioXml
[System.IO.File]::WriteAllText(
    $svgPath,
    $complexSvg,
    (New-Object System.Text.UTF8Encoding($false)))

$application = $null
$document = $null
$managedShape = $null
$plainShape = $null
$reopenedManaged = $null
$wordAddIn = $null
$measurementWordAddIn = $null
$selectionCounter = $null
$activeIdentity = $null
$reopenIdentity = $null
$builderIdentity = $null
$ownedIdentities =
    New-Object System.Collections.Generic.List[object]
$report = $null

try {
    Write-Host "Stage=BuildDocument"
    $buildResult = New-TypedStressDocument `
        -ImagePath $svgPath `
        -DocumentPath $documentPath `
        -DrawioXml $drawioXml `
        -CoreAssemblyPath $coreAssemblyPath `
        -WordAddInAssemblyPath $wordAddInAssemblyPath
    $builderIdentity = [pscustomobject]@{
        ProcessId = $buildResult.ProcessId
        StartTimeUtcTicks = $buildResult.ProcessStartTimeUtcTicks
    }
    $ownedIdentities.Add($builderIdentity)
    if (-not (Wait-WordProcessExit -Identity $builderIdentity)) {
        throw "The typed document-builder Word process did not exit."
    }

    Write-Host "Stage=StartVisibleWord"
    $application = New-Object -ComObject Word.Application
    $activeIdentity = Get-SoleWordProcessIdentity
    $ownedIdentities.Add($activeIdentity)
    $application.Visible = $true
    $application.DisplayAlerts = 0
    $application.ScreenUpdating = $true
    $document = Open-WordDocument `
        -Application $application `
        -Path $documentPath
    $activeWindowIdentity =
        Get-WordProcessIdentity -Application $application
    Assert-WordProcessIdentityMatch `
        -Expected $activeIdentity `
        -Actual $activeWindowIdentity
    $managedShape = Get-ShapeByName `
        -Document $document `
        -Name "Managed Complex v1.0.8"
    $plainShape = Get-ShapeByName `
        -Document $document `
        -Name "Plain Complex Same Visual"
    if ($null -eq $managedShape -or $null -eq $plainShape) {
        throw "The stress-test pictures were not found in the generated document."
    }
    $serializer = New-Object DrawioPpt.Core.Services.DiagramEnvelopeSerializer
    $measurementWordAddIn = Get-WordComAddIn `
        -Application $application `
        -ProgId "Greensoft.DrawioWordAddIn"
    $measurementWordAddInConnected =
        [bool]$measurementWordAddIn.Connect
    if (-not $measurementWordAddInConnected) {
        throw "The DrawioPpt Word add-in is not connected in the measured Word instance."
    }

    $selectionCounter =
        New-Object DrawioPptWordSelectionChangeCounter
    $selectionCounter.Attach($application)
    $selectionCountBefore = $selectionCounter.Count

    Write-Host "Stage=Round1PlainThenManaged"
    $round1Plain = Measure-ShapeMotion `
        -Document $document `
        -Shape $plainShape `
        -OtherShape $managedShape `
        -Label "Round1PlainComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity
    $round1Managed = Measure-ShapeMotion `
        -Document $document `
        -Shape $managedShape `
        -OtherShape $plainShape `
        -Label "Round1ManagedComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity

    Write-Host "Stage=Round2ManagedThenPlain"
    $round2Managed = Measure-ShapeMotion `
        -Document $document `
        -Shape $managedShape `
        -OtherShape $plainShape `
        -Label "Round2ManagedComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity
    $round2Plain = Measure-ShapeMotion `
        -Document $document `
        -Shape $plainShape `
        -OtherShape $managedShape `
        -Label "Round2PlainComplex" `
        -Seconds $DurationSeconds `
        -WarmupSeconds $WarmupSeconds `
        -ResizeOperations $ResizeOperationsPerRound `
        -SelectionSettleMilliseconds $SelectionSettleMilliseconds `
        -SelectionEventTimeoutMilliseconds $SelectionEventTimeoutMilliseconds `
        -SelectionCounter $selectionCounter `
        -WordProcessIdentity $activeIdentity

    $plainRounds = @($round1Plain, $round2Plain)
    $managedRounds = @($round1Managed, $round2Managed)
    $allRoundResults = @(
        $round1Plain,
        $round1Managed,
        $round2Managed,
        $round2Plain)
    $plainResult = Merge-ShapeMotionResults `
        -Results $plainRounds `
        -Label "PlainComplex"
    $managedResult = Merge-ShapeMotionResults `
        -Results $managedRounds `
        -Label "ManagedComplex"
    $observedSelectionChangeEvents =
        $selectionCounter.Count - $selectionCountBefore
    $expectedSelectionChangeEvents =
        ($allRoundResults |
            Measure-Object -Property SelectionOperations -Sum).Sum
    $confirmedShapeSelectionEvents =
        ($allRoundResults |
            Measure-Object `
                -Property ConfirmedShapeSelectionEvents `
                -Sum).Sum
    $missingShapeSelectionEvents =
        ($allRoundResults |
            Measure-Object `
                -Property MissingShapeSelectionEvents `
                -Sum).Sum
    $selectionChangeEventsPassed =
        $confirmedShapeSelectionEvents -eq
            $expectedSelectionChangeEvents -and
        $missingShapeSelectionEvents -eq 0
    $selectionCounter.Detach()
    $selectionCounter = $null
    Write-Host "Stage=MotionMeasured"

    $managedShape.Left = [single]($managedResult.StartLeft + 24)
    $managedShape.Top = [single]($managedResult.StartTop + 12)
    $managedShape.Width = [single]($managedShape.Width + 10)
    $managedShape.Height = [single]($managedShape.Height + 6)
    $plainShape.Visible = -1
    $plainShape.Left = 45
    $plainShape.Top = 390
    $plainShape.Width = 330
    $plainShape.Height = 220
    $expectedLeft = [single]$managedShape.Left
    $expectedTop = [single]$managedShape.Top
    $expectedWidth = [single]$managedShape.Width
    $expectedHeight = [single]$managedShape.Height
    $document.Save()
    $document.Close(0)
    Release-ComObject -Value $measurementWordAddIn
    $measurementWordAddIn = $null
    Release-ComObject -Value $plainShape
    $plainShape = $null
    Release-ComObject -Value $managedShape
    $managedShape = $null
    Release-ComObject -Value $document
    $document = $null
    $application.Quit()
    Release-ComObject -Value $application
    $application = $null

    if (-not (Wait-WordProcessExit -Identity $activeIdentity)) {
        throw "The first test-owned Word process did not exit within the timeout."
    }

    Write-Host "Stage=ReopenWord"
    $application = New-Object -ComObject Word.Application
    $reopenIdentity = Get-SoleWordProcessIdentity
    $ownedIdentities.Add($reopenIdentity)
    $application.Visible = $false
    $application.DisplayAlerts = 0
    $document = Open-WordDocument `
        -Application $application `
        -Path $documentPath
    $reopenWindowIdentity =
        Get-WordProcessIdentity -Application $application
    Assert-WordProcessIdentityMatch `
        -Expected $reopenIdentity `
        -Actual $reopenWindowIdentity
    $reopenedManaged = Get-ShapeByName `
        -Document $document `
        -Name "Managed Complex v1.0.8"
    if ($null -eq $reopenedManaged) {
        throw "Managed shape was not found after reopening the document."
    }

    $referenceSerialized = [string]$reopenedManaged.AlternativeText
    $reopenedReference = $serializer.Deserialize($referenceSerialized)
    $storedEnvelope = $null
    $customXmlParts = $null
    try {
        $customXmlParts = $document.CustomXMLParts
        for ($index = 1; $index -le $customXmlParts.Count; $index++) {
            $candidatePart = $null
            try {
                $candidatePart = $customXmlParts.Item($index)
                $candidateXml = [string]$candidatePart.XML
            }
            finally {
                Release-ComObject -Value $candidatePart
            }

            if (-not $serializer.CanDeserialize($candidateXml)) {
                continue
            }

            $candidateEnvelope = $serializer.Deserialize($candidateXml)
            if ($candidateEnvelope.DiagramId -eq
                $reopenedReference.DiagramId) {
                $storedEnvelope = $candidateEnvelope
                break
            }
        }
    }
    finally {
        Release-ComObject -Value $customXmlParts
    }

    $wordAddIn = Get-WordComAddIn `
        -Application $application `
        -ProgId "Greensoft.DrawioWordAddIn"
    $wordAddInConnected = [bool]$wordAddIn.Connect
    $reopenPersistencePassed =
        [Math]::Abs([single]$reopenedManaged.Left - $expectedLeft) -lt 0.1 -and
        [Math]::Abs([single]$reopenedManaged.Top - $expectedTop) -lt 0.1 -and
        [Math]::Abs([single]$reopenedManaged.Width - $expectedWidth) -lt 0.1 -and
        [Math]::Abs([single]$reopenedManaged.Height - $expectedHeight) -lt 0.1
    $storedPayloadPassed =
        $null -ne $storedEnvelope -and
        [string]$storedEnvelope.DrawioXml -eq $drawioXml -and
        [string]$reopenedReference.DrawioXml -eq "" -and
        $storedEnvelope.DiagramId -eq $reopenedReference.DiagramId

    $p95Ratio = if ($plainResult.P95Ms -gt 0) {
        [Math]::Round($managedResult.P95Ms / $plainResult.P95Ms, 3)
    }
    else {
        0
    }
    $selectionP95Ratio =
        if ($plainResult.SelectionP95Ms -gt 0) {
            [Math]::Round(
                $managedResult.SelectionP95Ms /
                    $plainResult.SelectionP95Ms,
                3)
        }
        else {
            0
        }
    $selectionP95DeltaMs = [Math]::Round(
        $managedResult.SelectionP95Ms -
            $plainResult.SelectionP95Ms,
        3)
    $resizeP95Ratio = if ($plainResult.ResizeP95Ms -gt 0) {
        [Math]::Round(
            $managedResult.ResizeP95Ms / $plainResult.ResizeP95Ms,
            3)
    }
    else {
        0
    }
    $round1P95Ratio = if ($round1Plain.P95Ms -gt 0) {
        [Math]::Round(
            $round1Managed.P95Ms / $round1Plain.P95Ms,
            3)
    }
    else {
        0
    }
    $round2P95Ratio = if ($round2Plain.P95Ms -gt 0) {
        [Math]::Round(
            $round2Managed.P95Ms / $round2Plain.P95Ms,
            3)
    }
    else {
        0
    }
    $round1SelectionP95Ratio =
        if ($round1Plain.SelectionP95Ms -gt 0) {
            [Math]::Round(
                $round1Managed.SelectionP95Ms /
                    $round1Plain.SelectionP95Ms,
                3)
        }
        else {
            0
        }
    $round2SelectionP95Ratio =
        if ($round2Plain.SelectionP95Ms -gt 0) {
            [Math]::Round(
                $round2Managed.SelectionP95Ms /
                    $round2Plain.SelectionP95Ms,
                3)
        }
        else {
            0
        }
    $round1SelectionP95DeltaMs = [Math]::Round(
        $round1Managed.SelectionP95Ms -
            $round1Plain.SelectionP95Ms,
        3)
    $round2SelectionP95DeltaMs = [Math]::Round(
        $round2Managed.SelectionP95Ms -
            $round2Plain.SelectionP95Ms,
        3)
    $round1ResizeP95Ratio =
        if ($round1Plain.ResizeP95Ms -gt 0) {
            [Math]::Round(
                $round1Managed.ResizeP95Ms /
                    $round1Plain.ResizeP95Ms,
                3)
        }
        else {
            0
        }
    $round2ResizeP95Ratio =
        if ($round2Plain.ResizeP95Ms -gt 0) {
            [Math]::Round(
                $round2Managed.ResizeP95Ms /
                    $round2Plain.ResizeP95Ms,
                3)
        }
        else {
            0
        }
    $roundP95Ratios = @($round1P95Ratio, $round2P95Ratio)
    $roundSelectionP95Ratios =
        @(
            $round1SelectionP95Ratio,
            $round2SelectionP95Ratio)
    $roundSelectionP95DeltasMs =
        @(
            $round1SelectionP95DeltaMs,
            $round2SelectionP95DeltaMs)
    $roundResizeP95Ratios =
        @($round1ResizeP95Ratio, $round2ResizeP95Ratio)
    $allSetFailures =
        ($allRoundResults |
            Measure-Object -Property SetFailures -Sum).Sum
    $allSelectionFailures =
        ($allRoundResults |
            Measure-Object -Property SelectionFailures -Sum).Sum
    $allResizeSetFailures =
        ($allRoundResults |
            Measure-Object -Property ResizeSetFailures -Sum).Sum
    $allUnresponsiveSamples =
        ($allRoundResults |
            Measure-Object -Property UnresponsiveSamples -Sum).Sum
    $selectionComparisonPassed =
        $selectionP95Ratio -gt 0 -and
        (
            $selectionP95Ratio -le $MaximumP95Ratio -or
            $selectionP95DeltaMs -le
                $MaximumSelectionP95DeltaMs
        ) -and
        @(
            for ($selectionRoundIndex = 0;
                $selectionRoundIndex -lt
                    $roundSelectionP95Ratios.Count;
                $selectionRoundIndex++) {
                if (
                    $roundSelectionP95Ratios[
                        $selectionRoundIndex] -le 0 -or
                    (
                        $roundSelectionP95Ratios[
                            $selectionRoundIndex] -gt
                                $MaximumPerRoundP95Ratio -and
                        $roundSelectionP95DeltasMs[
                            $selectionRoundIndex] -gt
                                $MaximumSelectionP95DeltaMs
                    )
                ) {
                    $false
                }
            }).Count -eq 0
    $comparisonPassed =
        $p95Ratio -gt 0 -and
        $p95Ratio -le $MaximumP95Ratio -and
        $selectionComparisonPassed -and
        $resizeP95Ratio -gt 0 -and
        $resizeP95Ratio -le $MaximumP95Ratio -and
        @($roundP95Ratios |
            Where-Object {
                $_ -le 0 -or
                $_ -gt $MaximumPerRoundP95Ratio
            }).Count -eq 0 -and
        @($roundResizeP95Ratios |
            Where-Object {
                $_ -le 0 -or
                $_ -gt $MaximumPerRoundP95Ratio
            }).Count -eq 0 -and
        $allSetFailures -eq 0 -and
        $allSelectionFailures -eq 0 -and
        $allResizeSetFailures -eq 0 -and
        $allUnresponsiveSamples -eq 0
    $absoluteLatencyGatePassed =
        @($allRoundResults |
            Where-Object {
                $_.SelectionP95Ms -gt
                    $MaximumAbsoluteSelectionP95Ms -or
                $_.P95Ms -gt $MaximumAbsoluteP95Ms -or
                $_.ResizeP95Ms -gt
                    $MaximumAbsoluteResizeP95Ms
            }).Count -eq 0
    $minimumSamplesPassed =
        @($allRoundResults |
            Where-Object {
                $_.Operations -lt $MinimumOperationsPerRound -or
                $_.SelectionOperations -lt
                    ($_.Operations + $_.ResizeOperations) -or
                $_.ResizeOperations -lt
                    $ResizeOperationsPerRound
            }).Count -eq 0
    $noAdditionalDifferenceObserved =
        $comparisonPassed
    $storagePassed =
        $reopenPersistencePassed -and $storedPayloadPassed

    $report = [ordered]@{
        ReportVersion = 8
        TestedVersion = "v1.0.8"
        ExecutedUtc = [DateTime]::UtcNow.ToString("o")
        InstallRoot = $installRootResolved
        TestDocument = $documentPath
        TestWordProcessIdentity =
            "$($activeIdentity.ProcessId):$($activeIdentity.StartTimeUtcTicks)"
        ReopenWordProcessIdentity =
            "$($reopenIdentity.ProcessId):$($reopenIdentity.StartTimeUtcTicks)"
        DurationSecondsPerPicturePerRound = $DurationSeconds
        WarmupSecondsPerPicturePerRound = $WarmupSeconds
        ResizeOperationsPerPicturePerRound =
            $ResizeOperationsPerRound
        SelectionSettleMilliseconds =
            $SelectionSettleMilliseconds
        SelectionEventTimeoutMilliseconds =
            $SelectionEventTimeoutMilliseconds
        MeasurementOrder = @(
            "Round1:PlainThenManaged",
            "Round2:ManagedThenPlain")
        MaximumP95Ratio = $MaximumP95Ratio
        MaximumPerRoundP95Ratio = $MaximumPerRoundP95Ratio
        MaximumSelectionP95DeltaMs =
            $MaximumSelectionP95DeltaMs
        MaximumAbsoluteSelectionP95Ms =
            $MaximumAbsoluteSelectionP95Ms
        MaximumAbsoluteP95Ms = $MaximumAbsoluteP95Ms
        MaximumAbsoluteResizeP95Ms =
            $MaximumAbsoluteResizeP95Ms
        MinimumOperationsPerRound = $MinimumOperationsPerRound
        SvgElementCount = 120
        DrawioXmlChars = $drawioXml.Length
        AlternativeTextChars = $referenceSerialized.Length
        Rounds = @(
            [ordered]@{
                Round = 1
                Order = "PlainThenManaged"
                Plain = $round1Plain
                Managed = $round1Managed
                ManagedPlainSelectionP95Ratio =
                    $round1SelectionP95Ratio
                ManagedPlainSelectionP95DeltaMs =
                    $round1SelectionP95DeltaMs
                ManagedPlainP95Ratio = $round1P95Ratio
                ManagedPlainResizeP95Ratio =
                    $round1ResizeP95Ratio
            },
            [ordered]@{
                Round = 2
                Order = "ManagedThenPlain"
                Plain = $round2Plain
                Managed = $round2Managed
                ManagedPlainSelectionP95Ratio =
                    $round2SelectionP95Ratio
                ManagedPlainSelectionP95DeltaMs =
                    $round2SelectionP95DeltaMs
                ManagedPlainP95Ratio = $round2P95Ratio
                ManagedPlainResizeP95Ratio =
                    $round2ResizeP95Ratio
            })
        Plain = $plainResult
        Managed = $managedResult
        ManagedPlainSelectionP95Ratio =
            $selectionP95Ratio
        ManagedPlainSelectionP95DeltaMs =
            $selectionP95DeltaMs
        ManagedPlainP95Ratio = $p95Ratio
        ManagedPlainResizeP95Ratio = $resizeP95Ratio
        SelectionComparisonPassed =
            $selectionComparisonPassed
        ComparisonPassed = $comparisonPassed
        AbsoluteLatencyGatePassed =
            $absoluteLatencyGatePassed
        MinimumSamplesPassed = $minimumSamplesPassed
        MeasurementWordAddInConnect =
            $measurementWordAddInConnected
        ObservedSelectionChangeEvents =
            $observedSelectionChangeEvents
        ExpectedSelectionChangeEvents =
            $expectedSelectionChangeEvents
        ConfirmedShapeSelectionEvents =
            $confirmedShapeSelectionEvents
        MissingShapeSelectionEvents =
            $missingShapeSelectionEvents
        SelectionChangeEventsPassed =
            $selectionChangeEventsPassed
        NoAdditionalMetadataPathDifferenceObserved = $noAdditionalDifferenceObserved
        ReopenPersistencePassed = $reopenPersistencePassed
        StoredPayloadPassed = $storedPayloadPassed
        StoragePassed = $storagePassed
        ReopenedPositionAndSize = [ordered]@{
            Left = [single]$reopenedManaged.Left
            Top = [single]$reopenedManaged.Top
            Width = [single]$reopenedManaged.Width
            Height = [single]$reopenedManaged.Height
        }
        StoredDrawioXmlChars = if ($null -ne $storedEnvelope) {
            ([string]$storedEnvelope.DrawioXml).Length
        }
        else {
            0
        }
        WordAddInConnect = $wordAddInConnected
        PointerInputCovered = $false
        PointerInputBoundary = "This test updates Word Shape properties on the visible Word STA UI thread; it does not inject mouse pointer input."
    }

    $document.Close(0)
    Release-ComObject -Value $wordAddIn
    $wordAddIn = $null
    Release-ComObject -Value $reopenedManaged
    $reopenedManaged = $null
    Release-ComObject -Value $document
    $document = $null
    $application.Quit()
    Release-ComObject -Value $application
    $application = $null

    if (-not (Wait-WordProcessExit -Identity $reopenIdentity)) {
        throw "The reopened test-owned Word process did not exit within the timeout."
    }
    Write-Host "Stage=ReopenVerified"

    $report.CleanupResidualWinWord = @(
        Get-Process WINWORD -ErrorAction SilentlyContinue).Count -gt 0
    $report.OverallPassed =
        $report.ComparisonPassed -and
        $report.AbsoluteLatencyGatePassed -and
        $report.MinimumSamplesPassed -and
        $report.MeasurementWordAddInConnect -and
        $report.SelectionChangeEventsPassed -and
        $report.StoragePassed -and
        $report.WordAddInConnect -and
        -not $report.CleanupResidualWinWord

    $reportJson = $report | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText(
        $outputFullPath,
        $reportJson,
        (New-Object System.Text.UTF8Encoding($false)))

    Write-Host "WordUiThreadStressReport=$outputFullPath"
    Write-Host "DurationSecondsPerPicturePerRound=$DurationSeconds"
    Write-Host "PlainP95Ms=$($plainResult.P95Ms)"
    Write-Host "ManagedP95Ms=$($managedResult.P95Ms)"
    Write-Host "ManagedPlainSelectionP95Ratio=$selectionP95Ratio"
    Write-Host "ManagedPlainSelectionP95DeltaMs=$selectionP95DeltaMs"
    Write-Host "ManagedPlainP95Ratio=$p95Ratio"
    Write-Host "ManagedPlainResizeP95Ratio=$resizeP95Ratio"
    Write-Host "ComparisonPassed=$($report.ComparisonPassed)"
    Write-Host "AbsoluteLatencyGatePassed=$($report.AbsoluteLatencyGatePassed)"
    Write-Host "MinimumSamplesPassed=$($report.MinimumSamplesPassed)"
    Write-Host "MeasurementWordAddInConnect=$($report.MeasurementWordAddInConnect)"
    Write-Host "ConfirmedShapeSelectionEvents=$($report.ConfirmedShapeSelectionEvents)"
    Write-Host "MissingShapeSelectionEvents=$($report.MissingShapeSelectionEvents)"
    Write-Host "SelectionChangeEventsPassed=$($report.SelectionChangeEventsPassed)"
    Write-Host "NoAdditionalMetadataPathDifferenceObserved=$($report.NoAdditionalMetadataPathDifferenceObserved)"
    Write-Host "ReopenPersistencePassed=$($report.ReopenPersistencePassed)"
    Write-Host "StoredPayloadPassed=$($report.StoredPayloadPassed)"
    Write-Host "WordAddInConnect=$($report.WordAddInConnect)"
    Write-Host "PointerInputCovered=False"
    Write-Host "CleanupResidualWinWord=$($report.CleanupResidualWinWord)"
    Write-Host "OverallPassed=$($report.OverallPassed)"

    if (-not $report.OverallPassed) {
        throw "Word UI-thread stress acceptance failed. See $outputFullPath"
    }
}
finally {
    if ($null -ne $selectionCounter) {
        try {
            $selectionCounter.Dispose()
        }
        catch {
        }
    }

    Release-ComObject -Value $measurementWordAddIn
    Release-ComObject -Value $wordAddIn
    Release-ComObject -Value $reopenedManaged
    Release-ComObject -Value $plainShape
    Release-ComObject -Value $managedShape
    if ($null -ne $document) {
        try {
            $document.Close(0)
        }
        catch {
        }

        Release-ComObject -Value $document
    }

    if ($null -ne $application) {
        try {
            $application.Quit()
        }
        catch {
        }

        Release-ComObject -Value $application
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    foreach ($identity in $ownedIdentities) {
        if (-not (Wait-WordProcessExit `
                -Identity $identity `
                -TimeoutSeconds 2)) {
            Stop-OwnedWordProcess -Identity $identity
        }
    }
}
