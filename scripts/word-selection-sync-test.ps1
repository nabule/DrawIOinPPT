$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$addInHostPath = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Services\AddInHost.cs"
$selectionMonitorPath = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Word\SelectionMonitor.cs"
$connectPath = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Connect.cs"
$ribbonControllerPath = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Ribbon\RibbonController.cs"
$ribbonXmlPath = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Ribbon\RibbonXmlProvider.cs"
$addInHostSource = [System.IO.File]::ReadAllText($addInHostPath)
$selectionMonitorSource = [System.IO.File]::ReadAllText($selectionMonitorPath)
$connectSource = [System.IO.File]::ReadAllText($connectPath)
$ribbonControllerSource = [System.IO.File]::ReadAllText($ribbonControllerPath)
$ribbonXmlSource = [System.IO.File]::ReadAllText($ribbonXmlPath)
$violations = @()

function Require-Source {
    param(
        [string]$Source,
        [string]$Snippet,
        [string]$Requirement
    )

    if ($Source.IndexOf($Snippet, [System.StringComparison]::Ordinal) -lt 0) {
        $script:violations += "Missing ${Requirement}: $Snippet"
    }
}

function Forbid-Source {
    param(
        [string]$Source,
        [string]$Snippet,
        [string]$Requirement
    )

    if ($Source.IndexOf($Snippet, [System.StringComparison]::Ordinal) -ge 0) {
        $script:violations += "Violated ${Requirement}: still found $Snippet"
    }
}

Forbid-Source $addInHostSource "System.Windows.Forms.Timer" "Word selection handling must not use Windows Forms Timer"
Forbid-Source $addInHostSource "_selectionStateTimer" "Word selection handling must not retain a polling lifecycle"
Forbid-Source $addInHostSource "OnSelectionStateTimerTick" "Word selection handling must not retain a Timer Tick handler"
Forbid-Source $selectionMonitorSource "WindowSelectionChange" "Word native selection and drag hot paths must not subscribe to selection events"
Forbid-Source $selectionMonitorSource "SelectionChanged" "Word selection monitoring must not propagate passive selection changes to the business layer"
Forbid-Source $addInHostSource "_selectionMonitor.SelectionChanged += OnSelectionChanged;" "AddInHost must not subscribe to passive selection changes"
Forbid-Source $addInHostSource "private void OnSelectionChanged" "AddInHost must not retain a passive selection handler"
Forbid-Source $addInHostSource "ApplySelectionContext(ReadLiveSelectionContext(), false);" "Startup must not read the current picture or Draw.io metadata"
Forbid-Source $addInHostSource "AutoOpenOnSelection" "The Word host must not retain selection-triggered auto-open behavior"
Forbid-Source $ribbonControllerSource "AutoOpen" "The Word Ribbon controller must not retain an auto-open callback"
Forbid-Source $connectSource "AutoOpen" "The Word COM entry point must not expose an auto-open callback"
Require-Source $connectSource "_comAddIn.Object = _automation;" "The installed Word host must expose an explicit edit automation entry point for E2E calls"
Require-Source $connectSource "public void EditSelectedDiagram()" "The explicit edit automation entry point must reuse the Ribbon edit command"
Require-Source $connectSource "_dispatcher.Invoke(new Action(_host.EditSelectedDiagram));" "The installed explicit-edit E2E path must dispatch to the Word STA UI thread"
Forbid-Source $ribbonXmlSource "Greensoft.DrawioWord.AutoOpen" "Word must not provide a selection-triggered auto-open entry point"
Require-Source $addInHostSource "SelectionContext selection = ReadLiveSelectionContext(out picture);" "An explicit operation must capture and validate the current picture once per click"
Require-Source $addInHostSource "new SettingsForm(_settings, _desktopEditorPathDetector, false)" "The Word settings window must hide the PowerPoint-only auto-open option"
Require-Source $ribbonControllerSource "return _host != null;" "Explicit-operation buttons must remain clickable and validate the selection when invoked"
$ribbonInstruction = -join @([char]0x9009, [char]0x4e2d, [char]0x56fe, [char]0x7247, [char]0x540e, [char]0x70b9, [char]0x51fb, [char]0x6b64, [char]0x6309, [char]0x94ae)
Require-Source $ribbonXmlSource $ribbonInstruction "The Ribbon tooltip must instruct users to select a picture before clicking"
Forbid-Source $addInHostSource "ResolveManagedPicture(" "An explicit operation must not fall back to a historical diagram ID after the live selection disappears"
Require-Source $addInHostSource "selection == null || !selection.HasSinglePicture || selection.IsManagedPicture" "The bind command must reject empty selections, multiple selections, and managed pictures"
Require-Source $addInHostSource "selection == null || !selection.IsManagedPicture" "The clear-binding command must reject regular pictures"

$liveSelectionReadCount = [regex]::Matches(
    $addInHostSource,
    [regex]::Escape("SelectionContext selection = ReadLiveSelectionContext(out picture);")).Count
if ($liveSelectionReadCount -lt 4) {
    $violations += "Explicit operations must each capture the live picture once: expected at least 4 occurrences, found $liveSelectionReadCount"
}

if ($violations.Count -gt 0) {
    $violations | ForEach-Object { Write-Error "Word zero-hot-path constraint failed: $_" }
    exit 1
}

Write-Host "WORD_SELECTION_SYNC_TEST_PASS: passive Word selection paths do not read metadata; explicit commands read the live selection after the click."
exit 0
