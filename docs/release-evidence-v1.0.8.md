# v1.0.8 发布证据与日志索引

[中文](./release-evidence-v1.0.8.md) | [English](./release-evidence-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 交付物

- 版本：`v1.0.8`
- 发布目录：`artifacts\releases\v1.0.8\package`
- 发布压缩包：`artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip`
- 标准本地安装根：`%LOCALAPPDATA%\Greensoft\DrawioPpt`
- 新增安装包回归：`word-svg-aspect-ratio-e2e.ps1`、`word-preview-image-provider-e2e.ps1`、`powerpoint-svg-aspect-ratio-e2e.ps1`

v1.0.8 仍处于待公开发布状态；本证据不表示 GitHub Release 已创建。

## 发布门槛

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-sync-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\webview2-user-data-folder-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-cleanup-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-preview-image-provider-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-svg-aspect-ratio-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\powerpoint-svg-aspect-ratio-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" -DrawioSourcePath "<USER_DRAWIO_SOURCE>" -PictureWrapMode Front -PictureRenderFormat Png -PreviewPixelWidth 620 -DurationSeconds 10 -MaximumP95Ratio 1.25 -MaximumPerRoundP95Ratio 1.25 -MaximumP95DeltaMs 50 -MaximumSelectionP95DeltaMs 50 -MaximumAbsoluteSelectionP95Ms 300 -MaximumAbsoluteP95Ms 300 -MaximumAbsoluteResizeP95Ms 300 -MinimumOperationsPerRound 10 -WarmupSeconds 4 -ResizeOperationsPerRound 12 -SelectionSettleMilliseconds 75 -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json"
```

## 最终验证结果

- `Release|x64` 构建通过：0 warnings，0 errors。
- 最终发布包的完整 Office 功能 E2E 为 `12/12 PASS`：`ReleasePackage`、`InstallRelease`、`InstalledOfficeAddInsVerify`、`PowerPointAddInLoad`、`InstalledUrlSmoke`、`InstalledWordUrlHostE2E`、`InstalledWordComplexMetadataE2E`、`InstalledWordSvgAspectE2E`、`InstalledWordPreviewProviderE2E`、`InstalledPowerPointSvgAspectE2E`、`PowerPointUrlE2E`、`DesktopExporter` 全部为 PASS。
- full E2E 汇总报告与日志副本内容一致，SHA256 为 `7D5AC38245C29AF08B04C24414480E1024E270210657EC0AFB97D0ABF52CF73C`；transcript SHA256 为 `3006C8C50E311F8F5BB65C9745BE131A440215FFB3B5A0780F96B1D773A37121`。
- 压缩阶段的包内 Markdown 链接门槛输出 `PackageMarkdownMissingLinkCount=0`。
- full E2E 只按 PID、启动时间和进程名识别测试拥有的 Office 进程，不终止用户预存进程。它在安装前快照 Word / PowerPoint 注册树，finally 精确恢复值、类型、子键和原本不存在状态；残留 `WINWORD` / `POWERPNT` PID 会写入 FatalError 并令 E2E 失败。
- 最终临时包卸载、注册状态恢复和设置恢复成功，输出 `FullE2ECleanupSucceeded=True`，没有 Office 进程残留。
- 最新源码 Word 正常展示为 620px 等比例 PNG 和 `wdWrapFront` 浮动图片；横向/纵向图均 `contain`。增强后的真实回归输出 `VerticalRatio=0.25`、`VerticalContained=True`、`FailedReplacementPreservedOriginal=True`，验证先建新图、成功后才删原图。导出失败生成 620×310 等比例轻量 PNG 占位预览，不回退复杂 SVG；临时 XML 使用重试、延迟和启动清理。先选中、再点击显式命令才读取选区并编辑。
- PowerPoint 直接插入原始 SVG，在幻灯片边界内按源比例 `contain` 居中，不改写 `viewBox`，不制造图内留白。

## DLL SHA256 与 ZIP 记录边界

以下哈希于 2026-07-29 在最终打包和标准本地安装后重新只读计算，三个 DLL 均在源码 Release、发布包和标准本地安装之间一致。

| 文件 | 源码 Release | 发布包 | 标准本地安装 |
| --- | --- | --- | --- |
| `DrawioPpt.Core.dll` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` |
| `DrawioPpt.PowerPointAddIn.dll` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` |
| `DrawioPpt.WordAddIn.dll` | `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D` | `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D` | `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D` |

相对位置：

- 源码 Release：`src\<project>\bin\x64\Release\<dll>`
- 发布包：`artifacts\releases\v1.0.8\package\bin\<dll>`
- 标准本地安装：`%LOCALAPPDATA%\Greensoft\DrawioPpt\bin\<dll>`

最终发布包已经包含 `6E6D…127D` Word DLL，并已覆盖安装到标准本地目录；安装后的 Word、预览提供器和 PowerPoint 比例 E2E 再次通过。

ZIP SHA256 不写入会再次进入同一 ZIP 的本证据文档，避免自引用导致哈希必然变化；由发布脚本或包外交付摘要在最终打包完成后记录。

## Word UI 线程压力证据

- 原始结果：`artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json`，`ReportVersion=9`，SHA256 `935D7C0827E872B5F29309EC1A30E0FB72898865CB4092BDBDAFE6B35928DCB1`。
- 测试图片：620×876 PNG、保持源比例、零边框、`Front` 浮动布局；普通图与受管图尺寸、锚点和环绕语义一致。
- 普通/受管位置 P95：`524.962/488.881 ms`；受管减普通：`-36.081 ms`；目标选区事件 `98/98`，缺失 0。
- 绝对位置 P95 门槛 `300 ms` 未通过；17×24 纯色 PNG 基线也未通过。
- `SelectionComparisonPassed=false`、`ComparisonPassed=false`、`AbsoluteLatencyGatePassed=false`、`NoAdditionalMetadataPathDifferenceObserved=false`、`OverallPassed=false`。
- `PointerInputCovered=false`；Windows `GetCursorPos` 访问被拒绝，`SetIsBorderRequired` 不受支持。本证据不声称鼠标拖动通过。
- 发布用[脱敏机器可读摘要](./evidence/word-ui-thread-stress-v1.0.8.json)已移除用户名、绝对附件路径、临时 GUID 和进程身份。

## 报告与日志

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json`
- [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md)
- [Word UI 线程压力测试报告](./word-ui-thread-stress-report-v1.0.8.md)
- [Word UI 线程压力脱敏证据](./evidence/word-ui-thread-stress-v1.0.8.json)

## 已知边界

最终包、标准本地安装和最新源码 DLL 已同步，功能性 `12/12 PASS` 已覆盖本次 Word 增强。独立压力测试仍因严格绝对 `300 ms` 门槛和真实指针未覆盖而为 `OverallPassed=false`；用户仍需完成真实指针拖动、缩放、重定位和人工点击“重新编辑”。最终 ZIP 哈希在包外交付摘要记录。
