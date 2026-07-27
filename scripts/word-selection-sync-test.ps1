$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$addInHostPath = Join-Path $repoRoot "src\DrawioPpt.WordAddIn\Services\AddInHost.cs"
$source = [System.IO.File]::ReadAllText($addInHostPath)
$violations = @()

function Require-Source {
    param(
        [string]$Snippet,
        [string]$Requirement
    )

    if ($source.IndexOf($Snippet, [System.StringComparison]::Ordinal) -lt 0) {
        $script:violations += "缺少 $Requirement：$Snippet"
    }
}

function Forbid-Source {
    param(
        [string]$Snippet,
        [string]$Requirement
    )

    if ($source.IndexOf($Snippet, [System.StringComparison]::Ordinal) -ge 0) {
        $script:violations += "违反 $Requirement：仍发现 $Snippet"
    }
}

Forbid-Source "System.Windows.Forms.Timer" "Word 选区同步不得使用 Windows Forms Timer"
Forbid-Source "Timer _selectionStateTimer" "Word 选区同步不得保留 Timer 字段"
Forbid-Source "_selectionStateTimer" "Word 选区同步不得保留 500ms 轮询生命周期"
Forbid-Source "OnSelectionStateTimerTick" "Word 选区同步不得保留 Timer Tick 处理器"
Require-Source "_selectionMonitor.Start();" "启动时必须启用 Word 选区事件监控"
Require-Source "ApplySelectionContext(ReadLiveSelectionContext(), false);" "启动时必须进行一次非自动打开的选区同步"

if ($violations.Count -gt 0) {
    $violations | ForEach-Object { Write-Error "Word 选区同步约束失败：$_" }
    exit 1
}

Write-Host "WORD_SELECTION_SYNC_TEST_PASS: Word 选区同步仅依赖事件与启动时单次同步。"
exit 0
