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
        $script:violations += "缺少 $Requirement：$Snippet"
    }
}

function Forbid-Source {
    param(
        [string]$Source,
        [string]$Snippet,
        [string]$Requirement
    )

    if ($Source.IndexOf($Snippet, [System.StringComparison]::Ordinal) -ge 0) {
        $script:violations += "违反 $Requirement：仍发现 $Snippet"
    }
}

Forbid-Source $addInHostSource "System.Windows.Forms.Timer" "Word 选区处理不得使用 Windows Forms Timer"
Forbid-Source $addInHostSource "_selectionStateTimer" "Word 选区处理不得保留轮询生命周期"
Forbid-Source $addInHostSource "OnSelectionStateTimerTick" "Word 选区处理不得保留 Timer Tick 处理器"
Forbid-Source $selectionMonitorSource "WindowSelectionChange" "Word 原生选中和拖动热路径不得订阅选区变化事件"
Forbid-Source $selectionMonitorSource "SelectionChanged" "Word 选区监控器不得向业务层传播被动选区变化"
Forbid-Source $addInHostSource "_selectionMonitor.SelectionChanged += OnSelectionChanged;" "AddInHost 不得订阅被动选区变化"
Forbid-Source $addInHostSource "private void OnSelectionChanged" "AddInHost 不得保留被动选区处理器"
Forbid-Source $addInHostSource "ApplySelectionContext(ReadLiveSelectionContext(), false);" "启动时不得读取当前图片或 Draw.io 元数据"
Forbid-Source $addInHostSource "AutoOpenOnSelection" "Word 宿主不得保留选中后自动打开逻辑"
Forbid-Source $ribbonControllerSource "AutoOpen" "Word Ribbon 控制器不得保留自动打开回调"
Forbid-Source $connectSource "AutoOpen" "Word COM 入口不得暴露自动打开回调"
Require-Source $connectSource "_comAddIn.Object = _automation;" "真实 Word 宿主必须公开显式编辑自动化入口供安装态 E2E 调用"
Require-Source $connectSource "public void EditSelectedDiagram()" "显式编辑自动化入口必须复用 Ribbon 编辑命令"
Require-Source $connectSource "_dispatcher.Invoke(new Action(_host.EditSelectedDiagram));" "安装态显式编辑 E2E 必须调度到 Word STA UI 线程"
Forbid-Source $ribbonXmlSource "Greensoft.DrawioWord.AutoOpen" "Word 不得提供选中后自动打开入口"
Require-Source $addInHostSource "SelectionContext selection = ReadLiveSelectionContext(out picture);" "显式操作必须一次捕获并校验点击时的图片快照"
Require-Source $addInHostSource "new SettingsForm(_settings, _desktopEditorPathDetector, false)" "Word 设置窗口必须隐藏仅 PowerPoint 使用的自动打开选项"
Require-Source $ribbonControllerSource "return _host != null;" "显式操作按钮必须保持可点击并在执行时校验选区"
Require-Source $ribbonXmlSource "选中图片后点击此按钮" "功能区提示必须说明先选中、再点击的交互"
Forbid-Source $addInHostSource "ResolveManagedPicture(" "显式操作不得在实时选区消失后按历史 diagramId 回退查找图片"
Require-Source $addInHostSource "selection == null || !selection.HasSinglePicture || selection.IsManagedPicture" "绑定命令必须拒绝空选区、多选和已受管图片"
Require-Source $addInHostSource "selection == null || !selection.IsManagedPicture" "清除绑定命令必须拒绝普通图片"

$liveSelectionReadCount = [regex]::Matches(
    $addInHostSource,
    [regex]::Escape("SelectionContext selection = ReadLiveSelectionContext(out picture);")).Count
if ($liveSelectionReadCount -lt 4) {
    $violations += "显式操作必须分别一次捕获实时图片：期望至少 4 处，实际 $liveSelectionReadCount 处"
}

if ($violations.Count -gt 0) {
    $violations | ForEach-Object { Write-Error "Word 零热路径约束失败：$_" }
    exit 1
}

Write-Host "WORD_SELECTION_SYNC_TEST_PASS: Word 被动选区路径不读取元数据，显式命令点击后读取实时选区。"
exit 0
