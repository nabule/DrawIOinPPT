# Word 与 PPT SVG 无留白及拖动性能 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Word 与 PowerPoint 新建、重新编辑和刷新 Draw.io SVG 时保持源图比例且不产生画布留白，并使用用户提供的 Draw.io 源码在包含正文、表格和其他图片的真实 Word 文档中验证拖动与缩放无插件热路径卡顿。

**Architecture:** 两个 Office 宿主直接插入原始 SVG，宽度保持现有布局语义，高度由源 SVG 比例计算；替换时保持旧图片宽度和垂直中心。性能验收复用 Word 零被动选区架构，以用户附件导出的 SVG 和原始 XML构建非空白文档，执行 COM 延迟对照及真实 Windows 指针拖动/缩放。

**Tech Stack:** C# 7.3、.NET Framework 4.8、Microsoft Office Word/PowerPoint COM Interop、PowerShell 5.1、draw.io Desktop、Windows UI Automation、MSBuild。

---

## 文件结构与职责

- `src/DrawioPpt.WordAddIn/Services/WordSvgPictureService.cs`：Word 新建和替换 SVG 的源比例几何。
- `src/DrawioPpt.PowerPointAddIn/Services/PowerPointSvgShapeService.cs`：PowerPoint 新建和替换 SVG 的源比例几何及元数据恢复。
- `scripts/word-svg-aspect-ratio-e2e.ps1`：真实 Word 2:1 SVG 新建与旧 1:1 图片修复回归。
- `scripts/powerpoint-svg-aspect-ratio-e2e.ps1`：真实 PowerPoint 2:1 SVG 新建与旧 1:1 图片修复回归。
- `scripts/word-ui-thread-stress-acceptance.ps1`：用户 Draw.io 源码、非空白 Word 文档、移动/缩放延迟和持久化验收。
- `scripts/package-release.ps1`：把两端比例回归和压力脚本纳入 v1.0.8。
- `scripts/full-e2e-test.ps1`：对临时安装的发布 DLL 串行执行 Word/PPT 比例回归。
- `docs/architecture*.md`、`docs/word-addin*.md`、`docs/user-guide*.md`、`docs/regression-checklist*.md`：两端几何规则、操作和人工验收。
- `docs/release-notes-v1.0.8*.md`、`docs/e2e-test-report-v1.0.8*.md`、`docs/release-evidence-v1.0.8*.md`、`docs/word-ui-thread-stress-report-v1.0.8*.md`：发布和实测证据。

用户附件 `C:\Users\nabul\.codex\attachments\0c8710b5-e3fb-4c0f-9267-bcdeabcd411d\pasted-text.txt` 只在本机测试时读取，不复制、不暂存、不打包。

### Task 1: 完成 Word 源比例红绿灯

**Files:**
- Modify: `src/DrawioPpt.WordAddIn/Services/WordSvgPictureService.cs`
- Create and track: `scripts/word-svg-aspect-ratio-e2e.ps1`

- [ ] **Step 1: 保留已取得的真实 Word 红灯证据**

现有 1200×600 SVG 测试在修复前应输出：

```text
InsertedRatio=1.613
ReplacedRatio=1.613
InsertPreservesSourceAspect=False
ReplacePreservesSourceAspect=False
```

- [ ] **Step 2: 核对最小实现仅处理几何**

`InsertAtSelection` 直接插入 `svgFilePath`，使用：

```csharp
float width = ResolveDefaultWidth(document);
float sourceAspectRatio = TryReadSourceAspectRatio(svgFilePath);
float height = ResolveHeightForWidth(width, sourceAspectRatio);
```

`Replace` 捕获旧图后执行：

```csharp
float sourceAspectRatio = TryReadSourceAspectRatio(svgFilePath);
snapshot.ApplySourceAspectRatio(sourceAspectRatio, !existingPicture.IsInline);
```

源比例有效时锁定比例；浮动图用 `(oldHeight - newHeight) / 2` 修正 `Top`。

- [ ] **Step 3: 运行真实 Word 绿灯**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-svg-aspect-ratio-e2e.ps1 -SkipBuild
```

Expected: `InsertedRatio=2`、`ReplacedRatio=2`、`ReplacementRetainsWidth=True`、`ReplacementCorrectsHeight=True`、`WORD_SVG_ASPECT_RATIO_E2E_PASS`。

### Task 2: 为 PowerPoint 建立红灯并实现无留白

**Files:**
- Create: `scripts/powerpoint-svg-aspect-ratio-e2e.ps1`
- Modify: `src/DrawioPpt.PowerPointAddIn/Services/PowerPointSvgShapeService.cs`

- [ ] **Step 1: 写真实 PowerPoint 红灯**

测试程序创建空白演示文稿和 1200×600 SVG，调用 `InsertOnActiveSlide`；然后把图片设为 400×400 并调用 `Replace`。断言：

```csharp
bool insertPreservesSourceAspect = Math.Abs(inserted.Width / inserted.Height - 2f) < 0.02f;
bool replacementRetainsWidth = Math.Abs(replaced.Width - 400f) < 0.1f;
bool replacementCorrectsHeight = Math.Abs(replaced.Height - 200f) < 0.1f;
bool replacementKeepsVerticalCenter = Math.Abs(
    (replaced.Top + replaced.Height / 2f) - legacyCenterY) < 0.1f;
bool replacementKeepsLeft = Math.Abs(replaced.Left - 100f) < 0.1f;
bool replacementLocksAspect = replaced.LockAspectRatio == MsoTriState.msoTrue;
```

- [ ] **Step 2: 运行红灯**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\powerpoint-svg-aspect-ratio-e2e.ps1 -SkipBuild
```

Expected: exit code `1`；当前实现新建比例约为 2.339，替换后仍为 1:1，至少 `InsertPreservesSourceAspect=False` 和 `ReplacePreservesSourceAspect=False`。

- [ ] **Step 3: 写最小 PowerPoint 实现**

新增源比例读取：

```csharp
private float TryReadSourceAspectRatio(string svgFilePath)
{
    float sourceWidth;
    float sourceHeight;
    if (_svgSupportService == null ||
        !_svgSupportService.TryReadSize(svgFilePath, out sourceWidth, out sourceHeight) ||
        sourceWidth <= 0f ||
        sourceHeight <= 0f)
    {
        return 0f;
    }

    return sourceWidth / sourceHeight;
}
```

新建保持当前默认宽度并按比例计算高度：

```csharp
float width = slideWidth * 0.50f;
float sourceAspectRatio = TryReadSourceAspectRatio(svgFilePath);
float height = sourceAspectRatio > 0f
    ? width / sourceAspectRatio
    : slideHeight * 0.38f;
```

直接执行 `slide.Shapes.AddPicture(svgFilePath, ...)`。替换前执行：

```csharp
snapshot.ApplySourceAspectRatio(TryReadSourceAspectRatio(svgFilePath));
```

`ShapeSnapshot.ApplySourceAspectRatio` 保留宽度、按旧图垂直中心修正 `Top`，并把 `LockAspectRatio` 设为 `msoTrue`。删除仅由本次改动变成无用的 `PreparePresentationSvg`。

- [ ] **Step 4: 构建并运行 PowerPoint 绿灯**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\powerpoint-svg-aspect-ratio-e2e.ps1 -SkipBuild
```

Expected: 构建 `0 warnings, 0 errors`；插入和替换比例均为 `2`，宽度 `400`、高度 `200`、垂直中心保持，脚本输出 `POWERPOINT_SVG_ASPECT_RATIO_E2E_PASS`。

### Task 3: 用用户 Draw.io 源码构建非空白 Word 压力场景

**Files:**
- Modify: `scripts/word-ui-thread-stress-acceptance.ps1`

- [ ] **Step 1: 增加附件输入与非空白文档红灯约束**

增加参数：

```powershell
[string]$DrawioSourcePath = "",
[int]$MinimumBodyParagraphs = 100,
[int]$MinimumBodyCharacters = 10000,
[int]$MinimumPageCount = 8,
[int]$MinimumTableCount = 2,
[int]$MinimumAuxiliaryPictureCount = 3,
[switch]$KeepArtifacts
```

若指定源码，复制到测试临时目录的 `.drawio` 文件，用 `draw.io.exe --export --format svg --output` 生成 SVG，并把源 XML 原样写入 `Document.CustomXMLParts`。构建结果必须新增并断言：

```text
BodyParagraphCount>=100
BodyCharacterCount>=10000
PageCount>=8
TableCount>=2
AuxiliaryPictureCount>=3
DrawioSourceChars=26106
DrawioSourceSha256=AB6E4281F949B9F7F3B147B3DE271C3F9F3CA24D06593D9920BADFCD0F05097D
```

在扩展构建器前运行脚本，Expected RED: 缺少非空白文档统计或未使用指定源码，验收退出 `1`。

- [ ] **Step 2: 构建真实非空白 Word 文档**

在测试文档中写入至少 100 段、10,000 字正文，形成至少 8 页；加入至少两个表格、三张不带 Draw.io 元数据的辅助图片以及非空页眉/页脚。受管图和普通对照图仍使用同一附件导出 SVG、相同尺寸和浮动环绕方式。测试记录正文哈希，并在保存重开后再次核对正文、页数、表格、辅助图片及页眉/页脚保持。

- [ ] **Step 3: 执行两轮交叉顺序 COM 压力测试**

```powershell
$attachment = 'C:\Users\nabul\.codex\attachments\0c8710b5-e3fb-4c0f-9267-bcdeabcd411d\pasted-text.txt'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 `
  -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" `
  -DrawioSourcePath $attachment `
  -DurationSeconds 12 `
  -WarmupSeconds 4 `
  -OutputPath .\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json
```

Expected: 普通图与受管图均达到最小操作数；受管/普通 P95 比例、绝对移动/缩放 P95、Word 响应、零失败和保存重开门槛全部通过；`NoPassiveSelectionMetadataPath=True`。

- [ ] **Step 4: 使用 Windows 桌面自动化做真实指针验收**

打开上一步的可见 Word 文档，对受管图片执行三轮各 10 秒连续拖动和每轮 10 次角点缩放；对普通图片执行相同动作。每轮动作前后截图，确认最终位置/尺寸命中、Word 窗口持续响应，操作期间没有编辑器弹出或插件元数据工作。验收要求 Word 无“未响应”、动作失败为 0，受管图相对普通图的 P95 差值不超过 50ms且比例不超过 1.25；真实指针观察到的单次停顿不得超过 300ms。若真实指针动作受桌面权限阻断，报告必须记录实际错误，不得用 COM 指标冒充指针验收。

- [ ] **Step 5: 安全清理测试资源**

测试前发现任何既有 `WINWORD.EXE` 即拒绝运行；仅按 PID、启动时间和窗口句柄识别测试拥有的 Word。附件复制到 `%TEMP%\DrawioPpt\word-real-pointer-<guid>` 后核对原文件与副本 SHA256 一致。COM 对象逆序释放并等待测试 Word 自然退出；仅在超时后终止精确身份匹配的测试进程。除非传入 `-KeepArtifacts`，递归删除前必须确认解析后的临时根位于上述专用目录内。

### Task 4: 接入发布包和完整 Office E2E

**Files:**
- Modify: `scripts/package-release.ps1`
- Modify: `scripts/full-e2e-test.ps1`

- [ ] **Step 1: 打包 PowerPoint 比例回归**

复制并在 `PACKAGE.txt` 列出：

```text
scripts\powerpoint-svg-aspect-ratio-e2e.ps1
```

- [ ] **Step 2: 完整 E2E 增加安装态 PowerPoint 比例门槛**

在安装态 Word 比例回归后串行调用：

```powershell
$installedPowerPointSvgAspectE2E = Join-Path $installRoot "scripts\powerpoint-svg-aspect-ratio-e2e.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installedPowerPointSvgAspectE2E `
    -Configuration Release `
    -AssemblyRoot (Join-Path $installRoot "bin") `
    -SkipBuild
Add-Result -Name "InstalledPowerPointSvgAspectE2E" -Passed ($LASTEXITCODE -eq 0) -Detail $installedPowerPointSvgAspectE2E
```

- [ ] **Step 3: 运行完整 Office E2E**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

Expected: `ReleasePackage`、安装/注册、URL smoke、真实 Word URL 宿主、复杂元数据、Word 比例、PowerPoint 比例、PowerPoint URL 回写和 draw.io Desktop 导出全部 PASS；无测试拥有的 Office 进程残留。

- [ ] **Step 4: 生成可移植 ZIP 并独立验证**

停止使用会生成反斜杠条目的 `Compress-Archive`。用 `System.IO.Compression.ZipArchive` 遍历 `package` 目录，并把相对路径中的 `\` 统一替换为 `/` 后创建条目。随后独立检查：

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipPath = '.\artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip'
$archive = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $zipPath))
try {
    $backslashEntries = @($archive.Entries | Where-Object { $_.FullName.Contains('\') })
    $required = @(
        'bin/DrawioPpt.WordAddIn.dll',
        'bin/DrawioPpt.PowerPointAddIn.dll',
        'scripts/word-svg-aspect-ratio-e2e.ps1',
        'scripts/powerpoint-svg-aspect-ratio-e2e.ps1'
    )
    $names = @($archive.Entries.FullName)
    if ($backslashEntries.Count -ne 0 -or @($required | Where-Object { $_ -notin $names }).Count -ne 0) {
        throw 'Portable ZIP verification failed.'
    }
}
finally {
    $archive.Dispose()
}
```

Expected: 反斜杠条目数为 `0`，必需文件全部存在；把 ZIP 解压到新的临时目录后，三项 DLL SHA256 与 `package\bin` 一致。

### Task 5: 更新中英文文档和发布证据

**Files:**
- Modify: `docs/architecture.md`, `docs/architecture.en.md`
- Modify: `docs/word-addin.md`, `docs/word-addin.en.md`
- Modify: `docs/user-guide.md`, `docs/user-guide.en.md`
- Modify: `docs/regression-checklist.md`, `docs/regression-checklist.en.md`
- Modify: `docs/release-notes-v1.0.8.md`, `docs/release-notes-v1.0.8.en.md`
- Modify: `docs/e2e-test-report-v1.0.8.md`, `docs/e2e-test-report-v1.0.8.en.md`
- Modify: `docs/release-evidence-v1.0.8.md`, `docs/release-evidence-v1.0.8.en.md`
- Modify: `docs/word-ui-thread-stress-report-v1.0.8.md`, `docs/word-ui-thread-stress-report-v1.0.8.en.md`
- Modify: `docs/evidence/word-ui-thread-stress-v1.0.8.json`

- [ ] **Step 1: 更新架构和使用说明**

明确 Word/PPT 均直接插入源 SVG；新建按源比例，替换保留宽度并重算高度；浮动图片保持垂直中心；无法读取尺寸时回退旧几何但不扩展 SVG 画布。

- [ ] **Step 2: 据实记录附件压力测试**

报告记录附件 SHA256、字符数、图元数、Word 正文段落/表格/辅助图片数量、每轮移动和缩放 P95、样本数、指针验收截图或权限错误。不得把路径中的用户名写入发布包内报告，路径统一写为 `<USER_DRAWIO_SOURCE>`。

- [ ] **Step 3: 更新 E2E 计数、DLL 哈希和发布清单**

以最终 `artifacts\test-reports\full-e2e-v1.0.8.md` 为准更新项目数；重新计算三项 DLL SHA256。ZIP 哈希在最终打包后单独记录，避免自哈希循环。

### Task 6: 最终打包、标准本地安装和提交

**Files:**
- Verify: `artifacts/releases/v1.0.8/package`
- Verify: `%LOCALAPPDATA%\Greensoft\DrawioPpt`

- [ ] **Step 1: 重打最终发布包并安装标准目录**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.8 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.8\package\scripts\install-release.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt"
```

- [ ] **Step 2: 对标准安装再跑关键门槛**

```powershell
$installRoot = "$env:LOCALAPPDATA\Greensoft\DrawioPpt"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$installRoot\scripts\verify-office-install.ps1" -InstallRoot $installRoot
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$installRoot\scripts\word-svg-aspect-ratio-e2e.ps1" -AssemblyRoot "$installRoot\bin" -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$installRoot\scripts\powerpoint-svg-aspect-ratio-e2e.ps1" -AssemblyRoot "$installRoot\bin" -SkipBuild
```

- [ ] **Step 3: 最终验证和中文提交**

```powershell
git diff --check
git status --short
Get-FileHash .\artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip -Algorithm SHA256
Get-Process WINWORD,POWERPNT -ErrorAction SilentlyContinue
```

仅暂存本任务文件，不暂存 `.playwright-cli/`、`docs/onenote-feasibility.md`、`scripts/desktop-diagram-monitor-manual-mode-test.ps1` 或用户附件。提交信息必须使用中文，并按“比例修复、压力测试、发布证据”说明独立功能点。

## 完成标准

- Word 和 PPT 新建、重新编辑、刷新均不扩展 SVG `viewBox`，不留白、不拉伸、不裁剪。
- 2:1 SVG 在两端新建和修复 400×400 旧图后均为 2:1；旧图宽度和浮动图垂直中心保持。
- 用户 Draw.io 源码在包含正文、表格和其他图片的 Word 文档中完成移动、缩放、保存重开和显式编辑验收。
- Word 拖动/缩放期间没有插件被动选区或元数据路径；COM 压力门槛和真实指针验收均据实记录。
- Release x64 构建、完整 Office E2E、发布包、标准本地安装和已安装 DLL 回归全部通过。
- 中英文架构、设计、使用说明、回归清单、发布说明和测试证据与实现一致。
