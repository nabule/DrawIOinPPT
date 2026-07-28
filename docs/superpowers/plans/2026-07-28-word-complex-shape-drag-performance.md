# Word 复杂图形拖动性能优化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 Word 图片的 `AlternativeText` 从完整 Draw.io envelope 改为轻量引用，完整 XML 可靠保存在 `Document.CustomXMLParts`；同时彻底移除 Word 被动选区热路径，采用“先选中、再点击命令”的显式编辑交互，并发布、安装和真实验证 `v1.0.8 Release x64`。

**Architecture:** `WordPictureMetadataService` 负责生成轻量图片引用和完整失败回退；`AddInHost` 负责保证先写文档级主存储、成功后再瘦身图片元数据。Word 不订阅 `WindowSelectionChange`，选中和拖动期间不执行插件处理；显式按钮点击后才读取实时选区并通过 `diagramId` 读取文档级完整 envelope。旧版完整图片元数据在主存储写入成功后自动迁移。

**Tech Stack:** C# 7.3、.NET Framework 4.8、Microsoft Office Word COM Interop、PowerShell 5.1、MSBuild、WebView2、真实 Word Desktop E2E。

---

## 文件结构与职责

- `src/DrawioPpt.WordAddIn/Services/WordPictureMetadataService.cs`：创建、写入和读取 Word 图片轻量引用；提供主存储失败时的完整回退。
- `src/DrawioPpt.WordAddIn/Services/AddInHost.cs`：协调 `CustomXMLParts` 主存储与图片引用的提交顺序，执行旧文档迁移。
- `scripts/word-complex-metadata-e2e.ps1`：真实 Word 复杂元数据红绿灯、完整存储、轻量引用和旧数据迁移验证。
- `scripts/word-url-e2e.ps1`：通过 `diagramId -> CustomXMLParts` 验证 URL 编辑后的完整 XML。
- `scripts/word-url-addin-host-e2e.ps1`：在真实 `WINWORD.EXE` 加载项宿主中验证轻量引用后的 URL 回写。
- `scripts/package-release.ps1`：将复杂元数据测试和 `v1.0.8` 资料纳入发布包。
- `scripts/full-e2e-test.ps1`：对临时安装的发布 DLL 串行执行复杂元数据 E2E。
- `README.md`、`README.en.md`：当前版本、下载入口和 Word 元数据摘要。
- `docs/architecture.md`、`docs/architecture.en.md`：文档级主存储和图片轻量索引架构。
- `docs/word-addin.md`、`docs/word-addin.en.md`：Word 保存、读取、迁移和复制边界。
- `docs/user-guide.md`、`docs/user-guide.en.md`：用户行为、版本入口和故障边界。
- `docs/regression-checklist.md`、`docs/regression-checklist.en.md`：复杂图形拖动和元数据迁移人工回归项。
- `docs/release-notes-v1.0.8.md`、`docs/release-notes-v1.0.8.en.md`：本版本行为变化。
- `docs/e2e-test-report-v1.0.8.md`、`docs/e2e-test-report-v1.0.8.en.md`：实际测试环境、结果和覆盖范围。
- `docs/release-evidence-v1.0.8.md`、`docs/release-evidence-v1.0.8.en.md`：最终发布物、命令、日志和哈希索引。

真实 Word、PowerPoint、注册表和安装测试必须串行执行。文档编辑、源码审查和不启动 Office 的静态检查可以并行。

### Task 1: 建立复杂元数据红灯并实现轻量图片引用

**Files:**
- Modify and track: `scripts/word-complex-metadata-e2e.ps1`
- Modify: `src/DrawioPpt.WordAddIn/Services/WordPictureMetadataService.cs`

- [ ] **Step 1: 复跑现有真实 Word 红灯**

确认 Word 未运行，然后直接使用当前 Release DLL：

```powershell
Get-Process WINWORD -ErrorAction SilentlyContinue
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
```

Expected: exit code `1`，输出包含：

```text
DrawioXmlChars=271361
AlternativeTextChars=<大于 2048>
StoredPayload=True
LightweightAlternativeText=False
SameDiagramId=True
```

- [ ] **Step 2: 在元数据服务中构造轻量 envelope**

在 `WordPictureMetadataService` 中将常规 `Save` 改为写入轻量副本，不修改调用方传入的完整对象：

```csharp
public void Save(WordPictureReference picture, DiagramEnvelope envelope)
{
    ValidateSaveArguments(picture, envelope);
    SaveEnvelope(picture, CreatePictureReferenceEnvelope(envelope));
}

private static DiagramEnvelope CreatePictureReferenceEnvelope(DiagramEnvelope envelope)
{
    DiagramEnvelope reference = new DiagramEnvelope();
    reference.FormatVersion = envelope.FormatVersion;
    reference.DiagramId = envelope.DiagramId;
    reference.DiagramName = envelope.DiagramName;
    reference.EditorMode = envelope.EditorMode;
    reference.EditorTarget = envelope.EditorTarget;
    reference.SidecarPath = envelope.SidecarPath;
    reference.UpdatedUtc = envelope.UpdatedUtc;
    reference.DrawioXml = string.Empty;
    return reference;
}

private void SaveEnvelope(WordPictureReference picture, DiagramEnvelope envelope)
{
    picture.Title = envelope.DiagramName ?? string.Empty;
    picture.AlternativeText = _serializer.Serialize(envelope);
}

private static void ValidateSaveArguments(WordPictureReference picture, DiagramEnvelope envelope)
{
    if (picture == null)
    {
        throw new ArgumentNullException("picture");
    }

    if (envelope == null)
    {
        throw new ArgumentNullException("envelope");
    }
}
```

- [ ] **Step 3: 构建并运行复杂元数据绿灯**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
```

Expected: build 为 `0 warnings, 0 errors`；E2E exit code `0`，`AlternativeTextChars <= 2048`、`StoredPayload=True`、`LightweightAlternativeText=True`、`SameDiagramId=True`。

- [ ] **Step 4: 检查改动范围并提交**

```powershell
git diff --check
git diff -- src/DrawioPpt.WordAddIn/Services/WordPictureMetadataService.cs scripts/word-complex-metadata-e2e.ps1
git add -- src/DrawioPpt.WordAddIn/Services/WordPictureMetadataService.cs scripts/word-complex-metadata-e2e.ps1
git commit -m "修复：瘦身 Word 图片元数据" -m "- 图片 AlternativeText 仅保存不含 DrawioXml 的轻量 envelope`n- 完整源数据继续保存在 Document.CustomXMLParts`n- 新增真实 Word 复杂元数据红绿灯"
```

### Task 2: 增加写入失败保护和旧文档安全迁移

**Files:**
- Modify: `scripts/word-complex-metadata-e2e.ps1`
- Modify: `src/DrawioPpt.WordAddIn/Services/WordPictureMetadataService.cs`
- Modify: `src/DrawioPpt.WordAddIn/Services/AddInHost.cs`

- [ ] **Step 1: 扩展失败回退和迁移测试**

在测试程序中增加 `System.Reflection`，创建第二张旧版图片，并通过生产 `AddInHost.TryReadManagedEnvelope` 路径验证迁移：

```csharp
using System.Reflection;

private static WordInterop.InlineShape AddPicture(
    WordInterop.Document document,
    string imagePath,
    int position)
{
    object linkToFile = false;
    object saveWithDocument = true;
    object range = document.Range(position, position);
    return document.InlineShapes.AddPicture(
        imagePath,
        ref linkToFile,
        ref saveWithDocument,
        ref range);
}

private static bool TestLegacyMigration(
    WordInterop.Application application,
    DiagramEnvelopeSerializer serializer,
    DocumentDiagramStore store,
    WordPictureMetadataService metadata,
    WordPictureReference legacyPicture,
    DiagramEnvelope envelope)
{
    metadata.SaveFallback(legacyPicture, envelope);

    AddInHost host = new AddInHost(application);
    try
    {
        MethodInfo method = typeof(AddInHost).GetMethod(
            "TryReadManagedEnvelope",
            BindingFlags.Instance | BindingFlags.NonPublic);
        object[] arguments = new object[] { legacyPicture, null };
        bool read = method != null && (bool)method.Invoke(host, arguments);
        DiagramEnvelope migrated = arguments[1] as DiagramEnvelope;
        DiagramEnvelope pictureReference = serializer.Deserialize(legacyPicture.AlternativeText);
        DiagramEnvelope stored;
        bool storedPayload = store.TryRead(
            application.ActiveDocument,
            envelope.DiagramId,
            out stored);

        return read &&
            migrated != null &&
            storedPayload &&
            stored != null &&
            string.Equals(stored.DrawioXml, envelope.DrawioXml, StringComparison.Ordinal) &&
            pictureReference != null &&
            string.IsNullOrEmpty(pictureReference.DrawioXml);
    }
    finally
    {
        host.Dispose();
    }
}
```

同时直接验证完整回退入口：

```csharp
WordPictureReference fallbackPicture = WordPictureReference.FromInlineShape(
    AddPicture(document, args[0], document.Content.End - 1));
metadata.SaveFallback(fallbackPicture, envelope);
DiagramEnvelope fallbackEnvelope = serializer.Deserialize(fallbackPicture.AlternativeText);
bool fallbackRetainsPayload =
    fallbackEnvelope != null &&
    string.Equals(fallbackEnvelope.DrawioXml, drawioXml, StringComparison.Ordinal);

DiagramEnvelope legacyEnvelope = new DiagramEnvelope();
legacyEnvelope.DiagramId = Guid.NewGuid().ToString("N");
legacyEnvelope.DiagramName = "Legacy migration test";
legacyEnvelope.DrawioXml = drawioXml;
WordPictureReference legacyPicture = WordPictureReference.FromInlineShape(
    AddPicture(document, args[0], document.Content.End - 1));
bool legacyMigrationPassed = TestLegacyMigration(
    application,
    serializer,
    store,
    metadata,
    legacyPicture,
    legacyEnvelope);

Console.WriteLine("FallbackRetainsPayload=" + fallbackRetainsPayload);
Console.WriteLine("LegacyMigrationPassed=" + legacyMigrationPassed);
return storedPayload &&
    lightweightAlternativeText &&
    sameDiagramId &&
    fallbackRetainsPayload &&
    legacyMigrationPassed
        ? 0
        : 1;
```

把原有首张图片创建调用同步改为 `AddPicture(document, args[0], 0)`，避免保留旧的双参数重载。

Expected RED: 编译失败并提示 `WordPictureMetadataService` 不存在 `SaveFallback`。

- [ ] **Step 2: 实现完整回退入口**

在 `WordPictureMetadataService` 中增加：

```csharp
public void SaveFallback(WordPictureReference picture, DiagramEnvelope envelope)
{
    ValidateSaveArguments(picture, envelope);
    SaveEnvelope(picture, envelope);
}
```

- [ ] **Step 3: 让宿主检查文档级写入结果**

把 `AddInHost.SaveManagedEnvelope` 的保存尾部改为：

```csharp
string customXmlPartId = _documentDiagramStore.Upsert(document, envelope);
if (string.IsNullOrWhiteSpace(customXmlPartId))
{
    _traceLog.Info(
        "WordAddInHost",
        "Document metadata storage was unavailable; retained the full picture fallback for diagram " +
        (envelope.DiagramId ?? string.Empty) + ".");
    _pictureMetadataService.SaveFallback(picture, envelope);
    return;
}

_pictureMetadataService.Save(picture, envelope);
```

`TryReadManagedEnvelope` 从 `CustomXMLParts` 读取成功后的 `_pictureMetadataService.Save(picture, storedEnvelope)` 保持不变；由于 `Save` 已变为轻量写入，该路径不会重新膨胀图片元数据。

- [ ] **Step 4: 验证迁移、回退和现有选区识别**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
```

Expected: 复杂测试输出 `FallbackRetainsPayload=True`、`LegacyMigrationPassed=True`；选区测试输出 `NoSelectionPolling=True`、`ManagedSelectionDetected=True`、`PlainPictureCanBind=True`。

- [ ] **Step 5: 提交失败保护和迁移**

```powershell
git diff --check
git add -- src/DrawioPpt.WordAddIn/Services/WordPictureMetadataService.cs src/DrawioPpt.WordAddIn/Services/AddInHost.cs scripts/word-complex-metadata-e2e.ps1
git commit -m "修复：保护 Word 元数据迁移与失败回退" -m "- 仅在 CustomXMLParts 写入成功后瘦身图片引用`n- 文档级写入失败时保留完整 DrawioXml`n- 验证旧版完整 AlternativeText 自动迁移"
```

### Task 3: 修复 Word URL 回归测试的完整数据读取方式

**Files:**
- Modify: `scripts/word-url-e2e.ps1`
- Modify: `scripts/word-url-addin-host-e2e.ps1`

- [ ] **Step 1: 证明旧测试假设已经失效**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -Configuration Release -SkipBuild
```

Expected RED: URL 流程可以保存，但 `HasDiagramState` 因图片引用的 `DrawioXml` 为空而返回失败。

- [ ] **Step 2: 让普通 Word URL E2E 从主存储读取**

将 `ReadEnvelope` 改为接收文档并解析主存储：

```csharp
private static DiagramEnvelope ReadEnvelope(
    WordInterop.Document document,
    WordInterop.InlineShape shape,
    DiagramEnvelopeSerializer serializer)
{
    if (document == null || shape == null || serializer == null)
    {
        return null;
    }

    string alternativeText = shape.AlternativeText ?? string.Empty;
    if (!serializer.CanDeserialize(alternativeText))
    {
        return null;
    }

    DiagramEnvelope reference = serializer.Deserialize(alternativeText);
    if (reference == null || string.IsNullOrWhiteSpace(reference.DiagramId))
    {
        return null;
    }

    DiagramEnvelope stored;
    return new DocumentDiagramStore(serializer).TryRead(document, reference.DiagramId, out stored)
        ? stored
        : reference;
}
```

调用点改为：

```csharp
DiagramEnvelope envelope = ReadEnvelope(document, document.InlineShapes[1], serializer);
```

- [ ] **Step 3: 让真实 Word 宿主测试扫描 `CustomXMLParts`**

在嵌入的 C# 程序中增加：

```csharp
private static DiagramEnvelope ReadStoredEnvelope(
    WordInterop.Document document,
    DiagramEnvelopeSerializer serializer,
    string diagramId)
{
    if (document == null || serializer == null || string.IsNullOrWhiteSpace(diagramId))
    {
        return null;
    }

    for (int index = 1; index <= document.CustomXMLParts.Count; index++)
    {
        string xml = document.CustomXMLParts[index].XML;
        if (!serializer.CanDeserialize(xml))
        {
            continue;
        }

        DiagramEnvelope candidate = serializer.Deserialize(xml);
        if (candidate != null &&
            string.Equals(candidate.DiagramId, diagramId, StringComparison.OrdinalIgnoreCase))
        {
            return candidate;
        }
    }

    return null;
}
```

将保存结果判断改为：

```csharp
DiagramEnvelopeSerializer serializer = new DiagramEnvelopeSerializer();
DiagramEnvelope reference = serializer.Deserialize(document.InlineShapes[1].AlternativeText);
DiagramEnvelope stored = ReadStoredEnvelope(document, serializer, reference.DiagramId);
string xml = stored == null ? string.Empty : stored.DrawioXml;
saved = !string.IsNullOrWhiteSpace(xml) &&
    xml.IndexOf("Round2", StringComparison.OrdinalIgnoreCase) >= 0;
```

- [ ] **Step 4: 串行运行两个真实 Word URL E2E**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release -SkipBuild
```

Expected: 普通 URL E2E 的新建、重开编辑、再次重开均为 `True`；安全约束输出 `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS`；真实宿主输出 `WordAddInConnect=True`、`ActualWordUrlEditorSaved=True`、`WebView2UserDataFolderExists=True`。

- [ ] **Step 5: 提交测试适配**

```powershell
git add -- scripts/word-url-e2e.ps1 scripts/word-url-addin-host-e2e.ps1
git commit -m "测试：按文档主存储验证 Word URL 回写" -m "- 通过图片 diagramId 读取 Document.CustomXMLParts`n- 保留新建、保存重开和真实 WINWORD 宿主回归`n- 移除测试对完整 AlternativeText 的旧假设"
```

### Task 4: 更新架构、使用说明和版本入口

**Files:**
- Modify: `README.md`
- Modify: `README.en.md`
- Modify: `docs/architecture.md`
- Modify: `docs/architecture.en.md`
- Modify: `docs/word-addin.md`
- Modify: `docs/word-addin.en.md`
- Modify: `docs/user-guide.md`
- Modify: `docs/user-guide.en.md`
- Modify: `docs/regression-checklist.md`
- Modify: `docs/regression-checklist.en.md`
- Create: `docs/release-notes-v1.0.8.md`
- Create: `docs/release-notes-v1.0.8.en.md`

- [ ] **Step 1: 更新架构数据模型**

中英文架构必须明确：

```text
Word 的 Document.CustomXMLParts 保存包含完整 DrawioXml 的主 envelope。
InlineShape/Shape.AlternativeText 只保存不含 DrawioXml 的轻量引用 envelope。
旧版完整 AlternativeText 在主存储写入成功后自动迁移；写入失败时保留完整回退。
单独复制图片到另一文档不再保证恢复完整 DrawioXml。
```

- [ ] **Step 2: 更新 Word 设计和用户说明**

将“`AlternativeText` 保存压缩后的完整 Draw.io 元数据包”改为：

```text
图片 AlternativeText 保存轻量引用 envelope，用于快速识别 diagramId；
完整 Draw.io XML 仅保存在 Document.CustomXMLParts。
```

在回归清单加入：

```markdown
- [ ] 插入 1,200 图元复杂图后，`AlternativeText <= 2048` 且不含 `DrawioXml`
- [ ] 复杂浮动图形连续拖动、缩放时不再出现插件引起的周期性停顿
- [ ] 旧版完整 `AlternativeText` 在打开编辑后迁移到 `CustomXMLParts`
- [ ] 文档级写入失败时仍保留完整图片回退
```

- [ ] **Step 3: 新增 v1.0.8 发布说明并更新版本入口**

发布说明只陈述已实现行为，不预写未执行的测试结果。把 README 和用户指南的当前版本、下载包和版本资料链接从 `v1.0.7` 更新为 `v1.0.8`；安装文档中“从 v1.0.7 起使用用户级 WebView2 目录”的历史说明保持不变。

- [ ] **Step 4: 校验中英文一致性并提交**

```powershell
rg -n "v1\\.0\\.7" README.md README.en.md docs/user-guide.md docs/user-guide.en.md
rg -n "完整.*AlternativeText|full.*AlternativeText|压缩后的 Draw.io 元数据包|compressed Draw.io metadata envelope" docs/architecture.md docs/architecture.en.md docs/word-addin.md docs/word-addin.en.md
git diff --check
git add -- README.md README.en.md docs/architecture.md docs/architecture.en.md docs/word-addin.md docs/word-addin.en.md docs/user-guide.md docs/user-guide.en.md docs/regression-checklist.md docs/regression-checklist.en.md docs/release-notes-v1.0.8.md docs/release-notes-v1.0.8.en.md
git commit -m "文档：说明 Word 轻量元数据存储策略" -m "- 更新架构、Word 设计、用户指南和回归清单`n- 明确旧文档迁移与单图片复制边界`n- 新增 v1.0.8 中英文发布说明和下载入口"
```

Expected: 前两个 `rg` 不得命中仍声称“当前 Word 图片保存完整回退”的文本；历史版本文档不修改。

### Task 5: 将复杂元数据回归接入 v1.0.8 发布链

**Files:**
- Modify: `scripts/package-release.ps1`
- Modify: `scripts/full-e2e-test.ps1`
- Create: `docs/e2e-test-report-v1.0.8.md`
- Create: `docs/e2e-test-report-v1.0.8.en.md`
- Create: `docs/release-evidence-v1.0.8.md`
- Create: `docs/release-evidence-v1.0.8.en.md`

- [ ] **Step 1: 记录打包红灯**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.8 -Configuration Release -Platform x64 -SkipBuild
```

Expected RED: 在 v1.0.8 测试报告和发布证据创建前，打包因缺少版本化资料失败；不得复制或改名 v1.0.7 资料冒充新结果。

- [ ] **Step 2: 更新默认版本和发布包脚本清单**

把两个脚本的默认版本改为：

```powershell
[string]$Version = "v1.0.8"
```

在 `package-release.ps1` 中复制并列出：

```powershell
Copy-Item (Join-Path $repoRoot "scripts\\word-complex-metadata-e2e.ps1") (Join-Path $packageRoot "scripts") -Force
```

`PACKAGE.txt` 增加：

```text
- scripts\word-complex-metadata-e2e.ps1
```

- [ ] **Step 3: 将已安装 DLL 的复杂测试加入完整 E2E**

在 `InstalledWordUrlHostE2E` 后串行加入：

```powershell
$installedWordComplexMetadataE2E = Join-Path $installRoot "scripts\\word-complex-metadata-e2e.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installedWordComplexMetadataE2E `
    -Configuration Release `
    -SkipBuild `
    -AssemblyRoot (Join-Path $installRoot "bin")
Add-Result `
    -Name "InstalledWordComplexMetadataE2E" `
    -Passed ($LASTEXITCODE -eq 0) `
    -Detail $installedWordComplexMetadataE2E
```

- [ ] **Step 4: 先运行独立门槛并据实创建版本报告**

串行运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-sync-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\webview2-user-data-folder-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release -SkipBuild
```

根据实际输出创建四份中英文报告；不得填写未运行的 `PASS`。报告至少记录 `AlternativeTextChars`、`StoredPayload`、`LightweightAlternativeText`、`FallbackRetainsPayload`、`LegacyMigrationPassed` 和 Word URL 结果。

- [ ] **Step 5: 验证打包和完整 E2E**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.8 -Configuration Release -Platform x64 -SkipBuild
Test-Path .\artifacts\releases\v1.0.8\package\scripts\word-complex-metadata-e2e.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

Expected: 复杂脚本存在；完整 E2E 新增 `InstalledWordComplexMetadataE2E=PASS`，总结果为 9/9 PASS。若实际项目数变化，以报告表格全部为 PASS 为准，不硬编码伪造计数。

- [ ] **Step 6: 用完整 E2E 结果更新报告并重新打最终包**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.8 -Configuration Release -Platform x64 -SkipBuild
Get-FileHash .\artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip -Algorithm SHA256
```

把完整 E2E 报告路径、最终项目结果和各 DLL SHA256 写入发布证据。压缩包本身包含发布证据，不能把自己的哈希写回包内；重新打包后计算最终 ZIP SHA256，并在任务交付消息和 All Notify 完成通知中记录。

- [ ] **Step 7: 提交发布链和实测报告**

```powershell
git add -- scripts/package-release.ps1 scripts/full-e2e-test.ps1 docs/e2e-test-report-v1.0.8.md docs/e2e-test-report-v1.0.8.en.md docs/release-evidence-v1.0.8.md docs/release-evidence-v1.0.8.en.md
git commit -m "发布：接入 v1.0.8 复杂元数据回归" -m "- 发布包新增真实 Word 复杂元数据测试`n- 完整 Office E2E 验证已安装 DLL 的轻量引用`n- 记录 v1.0.8 中英文实测报告和发布证据"
```

### Task 6: 安装本地 v1.0.8 并做真实 Word 拖动验收

**Files:**
- Verify: `artifacts/releases/v1.0.8/package/PACKAGE.txt`
- Verify: `artifacts/releases/v1.0.8/DrawioPpt-v1.0.8.zip`
- Verify installed root: `%LOCALAPPDATA%\Greensoft\DrawioPpt`
- Update if evidence changes: `docs/e2e-test-report-v1.0.8.md`
- Update if evidence changes: `docs/e2e-test-report-v1.0.8.en.md`
- Update if evidence changes: `docs/release-evidence-v1.0.8.md`
- Update if evidence changes: `docs/release-evidence-v1.0.8.en.md`

- [ ] **Step 1: 确认 Office 进程已关闭并安装标准本地目录**

```powershell
Get-Process WINWORD,POWERPNT -ErrorAction SilentlyContinue
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.8\package\scripts\install-release.ps1
```

Expected: 安装根目录为 `%LOCALAPPDATA%\Greensoft\DrawioPpt`，脚本输出 `DrawioPpt installed successfully.`。

- [ ] **Step 2: 验证本地注册、加载路径和 COM 加载**

```powershell
$installRoot = Join-Path $env:LOCALAPPDATA "Greensoft\\DrawioPpt"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $installRoot "scripts\\verify-office-install.ps1") -InstallRoot $installRoot
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $installRoot "scripts\\word-addin-load-check.ps1") `
    -Configuration Release `
    -AssemblyPath (Join-Path $installRoot "bin\\DrawioPpt.WordAddIn.dll") `
    -KeepRegistered
Get-Content -LiteralPath (Join-Path $installRoot "PACKAGE.txt")
```

Expected: PowerPoint 和 Word 注册路径都指向本地安装目录，Word 加载检查通过，`PACKAGE.txt` 显示 `Version: v1.0.8`。

- [ ] **Step 3: 对已安装 DLL 再跑关键 Word 回归**

```powershell
$installRoot = Join-Path $env:LOCALAPPDATA "Greensoft\\DrawioPpt"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $installRoot "scripts\\word-complex-metadata-e2e.ps1") -Configuration Release -SkipBuild -AssemblyRoot (Join-Path $installRoot "bin")
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $installRoot "scripts\\word-selection-event-e2e.ps1") -Configuration Release -SkipBuild -AssemblyRoot (Join-Path $installRoot "bin")
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $installRoot "scripts\\word-url-addin-host-e2e.ps1") -Configuration Release -SkipBuild -AssemblyRoot (Join-Path $installRoot "bin")
```

Expected: 三项真实 Word 回归全部通过，测试结束后没有遗留 `WINWORD.EXE`。

- [ ] **Step 4: 使用已安装插件做可见 Word 拖动对照**

在真实 Word 中使用已安装插件插入或绑定复杂图形，转换为浮动图片后执行：

```text
1. 连续拖动 10 秒。
2. 连续调整大小 10 次。
3. 在页面上方、正文中部和页面底部重新定位。
4. 与同尺寸普通图片执行相同操作。
5. 关闭并重新打开文档，再次编辑该 Draw.io 图形。
```

Expected: 复杂受管图形位置连续跟随指针，不再出现随元数据解析产生的周期性停顿；重新打开后仍能从 `CustomXMLParts` 恢复完整源 XML。若同一 SVG 在禁用插件后仍有等量卡顿，单独记录为 Word 原生 SVG/排版开销。

- [ ] **Step 5: 最终验证工作区和证据**

```powershell
git status --short
git log -6 --oneline
Get-FileHash .\artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip -Algorithm SHA256
Get-Process WINWORD,POWERPNT -ErrorAction SilentlyContinue
```

Expected: 仅保留用户原有的无关未跟踪文件；所有实现、测试和文档提交均为中文；无遗留 Office 测试进程；最终哈希与交付记录一致。

- [ ] **Step 6: 如人工验收补充了证据，提交最终记录**

仅在四份报告确有新增实测内容时执行：

```powershell
git add -- docs/e2e-test-report-v1.0.8.md docs/e2e-test-report-v1.0.8.en.md docs/release-evidence-v1.0.8.md docs/release-evidence-v1.0.8.en.md
git commit -m "测试：补充 v1.0.8 本地 Word 验收证据" -m "- 记录本地安装路径与 Word COM 加载结果`n- 补充复杂图形拖动、缩放和重开编辑验证`n- 固化最终 DLL 哈希和发布包外部哈希索引"
```

## 完成标准

- 所有 Word 新写入图片使用轻量 `AlternativeText`，完整 XML 只保存在 `Document.CustomXMLParts`。
- Word 发布 DLL 不含 `WindowSelectionChange` 订阅、被动选区回调或启动时选区/元数据读取；Word Ribbon 不提供自动打开入口。
- 用户先选中图片，再点击“重新编辑 / 刷新 / 绑定 / 清除绑定”；命令读取点击时的实时选区，不能对历史缓存对象执行操作。
- 文档级写入失败时保留完整图片回退，不丢失源数据。
- 旧版完整图片元数据可安全迁移。
- PowerPoint 行为和元数据策略不变。
- 自动化构建、复杂元数据、选区、URL、真实宿主、打包和完整 Office E2E 全部通过。
- `v1.0.8` 发布包已安装到当前用户标准目录，并通过真实 Word 拖动与重开编辑验收。
- 架构、设计、用户说明、回归清单、发布说明和测试证据与实现一致。
