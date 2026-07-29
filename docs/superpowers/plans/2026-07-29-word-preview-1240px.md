# Word 1240px Preview Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把 Word Draw.io PNG 预览从 620px 提升到 1240px，并在真实富文档中确认清晰度与交互性能。

**Architecture:** 只修改 `WordPreviewImageProvider` 的共享预览宽度常量，保留 PNG 浮动图片、`wdWrapFront`、显式编辑和原有清理机制。测试先锁定 1240px 正常与回退输出，再更新实现、文档、发布包和本地安装。

**Tech Stack:** C# .NET Framework 4.8、PowerShell E2E、Microsoft Word COM、draw.io Desktop

---

### Task 1: 预览宽度回归

**Files:**
- Modify: `scripts/word-preview-image-provider-e2e.ps1`
- Modify: `src/DrawioPpt.WordAddIn/Services/WordPreviewImageProvider.cs`

- [ ] **Step 1: 写入失败断言**

把正常 PNG 与轻量占位 PNG 的期望宽度改为 1240，2:1 占位图期望尺寸改为 1240×620，并把输出键改为 `Width1240Passed`。

- [ ] **Step 2: 验证测试先失败**

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-preview-image-provider-e2e.ps1 -Configuration Release -SkipBuild
```

Expected: 当前 620px 实现导致 `Width1240Passed=False` 并以非零状态退出。

- [ ] **Step 3: 最小实现**

把 `WordPreviewImageProvider.PreviewWidth` 从 `620` 改为 `1240`，不修改其他展示或交互逻辑。

- [ ] **Step 4: 构建并验证通过**

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-preview-image-provider-e2e.ps1 -Configuration Release -SkipBuild
```

Expected: 构建 0 错误、E2E 输出 `PixelWidth=1240`、`Width1240Passed=True` 和 `WORD_PREVIEW_IMAGE_PROVIDER_E2E_PASS`。

### Task 2: 文档一致性

**Files:**
- Modify: `docs/architecture.md`
- Modify: `docs/architecture.en.md`
- Modify: `docs/word-addin.md`
- Modify: `docs/word-addin.en.md`
- Modify: `docs/user-guide.md`
- Modify: `docs/user-guide.en.md`
- Modify: `docs/regression-checklist.md`
- Modify: `docs/regression-checklist.en.md`

- [ ] **Step 1: 更新当前设计与使用说明**

把当前正常预览与轻量占位图规则更新为 1240px，明确 PowerPoint 仍使用原始 SVG、Word 仍保留显式编辑和无被动元数据热路径。

- [ ] **Step 2: 精确检索陈旧当前态描述**

Run:

```powershell
rg -n "620px|620 px|620×310" docs\architecture*.md docs\word-addin*.md docs\user-guide*.md docs\release-notes-v1.0.8*.md
```

Expected: 当前态架构、实现、使用与回归契约不再把 620px 描述为最新实现；旧压力数据只在明确标注的历史样本中保留。

### Task 3: 真实富文档压力测试

**Files:**
- Modify: `scripts/word-ui-thread-stress-acceptance.ps1`
- Update: `docs/evidence/word-ui-thread-stress-v1.0.8.json`
- Update: `docs/word-ui-thread-stress-report-v1.0.8.md`
- Update: `docs/word-ui-thread-stress-report-v1.0.8.en.md`

- [ ] **Step 1: 把请求像素宽度纳入压力测试门禁**

压力脚本必须校验 PNG 实际宽度与 `PreviewPixelWidth` 相等，并把结果写入 `Rendering.PixelWidthPassed` 且并入 `Rendering.Passed`/`SourcePassed`；这样命令行参数不能冒充实际 1240px 输出。

- [ ] **Step 2: 使用用户复杂源图运行 1240px 富文档测试**

Run:

```powershell
$userSourcePath = Join-Path $env:USERPROFILE ".codex\attachments\0c8710b5-e3fb-4c0f-9267-bcdeabcd411d\pasted-text.txt"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" -DrawioSourcePath $userSourcePath -PictureWrapMode Front -PictureRenderFormat Png -PreviewPixelWidth 1240 -DurationSeconds 10 -MaximumP95Ratio 1.25 -MaximumPerRoundP95Ratio 1.25 -MaximumP95DeltaMs 50 -MaximumSelectionP95DeltaMs 50 -MaximumAbsoluteSelectionP95Ms 300 -MaximumAbsoluteP95Ms 300 -MaximumAbsoluteResizeP95Ms 300 -MinimumOperationsPerRound 10 -WarmupSeconds 4 -ResizeOperationsPerRound 12 -SelectionSettleMilliseconds 75 -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json"
```

Expected: 文档包含 150 个正文段落、表格和辅助图片，预览宽度 1240px、比例一致、无边框；记录普通图与受管图的选择、移动、缩放对照结果。

- [ ] **Step 3: 真实 Word 人工交互复核**

在可见 Word 中连续拖动、调整大小，并选中后点击“重新编辑”；记录鼠标验收事实，不用 COM 属性延迟冒充真实鼠标结论。

### Task 4: 发布与本地更新

**Files:**
- Update: `artifacts/releases/v1.0.8/package`
- Update: `artifacts/releases/v1.0.8/DrawioPpt-v1.0.8.zip`

- [ ] **Step 1: 运行完整回归与打包**

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.8 -Configuration Release -Platform x64 -SkipBuild
```

Expected: 完整回归无失败，包内 Word DLL 与最新 Release 构建一致。

- [ ] **Step 2: 更新本地安装并复核**

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.8\package\scripts\install-release.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.8\package\scripts\verify-office-install.ps1
```

Expected: 本地安装目录、发布包和最新源码 Word DLL 的 SHA-256 一致，Word 加载项可加载。
