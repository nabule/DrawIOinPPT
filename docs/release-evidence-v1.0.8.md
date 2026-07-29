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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" -DrawioSourcePath "<USER_DRAWIO_SOURCE>" -PictureWrapMode Front -PictureRenderFormat Png -PreviewPixelWidth 1240 -DurationSeconds 10 -MaximumP95Ratio 1.25 -MaximumPerRoundP95Ratio 1.25 -MaximumP95DeltaMs 50 -MaximumSelectionP95DeltaMs 50 -MaximumAbsoluteSelectionP95Ms 300 -MaximumAbsoluteP95Ms 300 -MaximumAbsoluteResizeP95Ms 300 -MinimumOperationsPerRound 10 -WarmupSeconds 4 -ResizeOperationsPerRound 12 -SelectionSettleMilliseconds 75 -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json"
```

## 最终验证结果

- `Release|x64` 构建通过：0 warnings，0 errors。
- 最终发布包的完整 Office 功能 E2E 为 `12/12 PASS`：`ReleasePackage`、`InstallRelease`、`InstalledOfficeAddInsVerify`、`PowerPointAddInLoad`、`InstalledUrlSmoke`、`InstalledWordUrlHostE2E`、`InstalledWordComplexMetadataE2E`、`InstalledWordSvgAspectE2E`、`InstalledWordPreviewProviderE2E`、`InstalledPowerPointSvgAspectE2E`、`PowerPointUrlE2E`、`DesktopExporter` 全部为 PASS。
- full E2E 汇总报告与日志副本内容一致，SHA256 为 `7D5AC38245C29AF08B04C24414480E1024E270210657EC0AFB97D0ABF52CF73C`；transcript SHA256 为 `0E629301AEE84FE71733B0FD956747201E99EADE70A2B3FE4E1DDBDBC55E4932`。
- 压缩阶段的包内 Markdown 链接门槛输出 `PackageMarkdownMissingLinkCount=0`。
- full E2E 只按 PID、启动时间和进程名识别测试拥有的 Office 进程，不终止用户预存进程。它在安装前快照 Word / PowerPoint 注册树，finally 精确恢复值、类型、子键和原本不存在状态；残留 `WINWORD` / `POWERPNT` PID 会写入 FatalError 并令 E2E 失败。
- 最终临时包卸载、注册状态恢复和设置恢复成功，输出 `FullE2ECleanupSucceeded=True`，没有 Office 进程残留。
- 最新源码 Word 正常展示为 1240px 等比例 PNG 和 `wdWrapFront` 浮动图片；横向/纵向图均 `contain`。增强后的真实回归输出 `VerticalRatio=0.25`、`VerticalContained=True`、`FailedReplacementPreservedOriginal=True`，验证先建新图、成功后才删原图。导出失败按 1240px 基准生成等比例轻量 PNG 占位预览，极端纵向图高度限制为 1200px，不回退复杂 SVG；临时 XML 实现包含重试、延迟和启动清理。先选中、再点击显式命令才读取选区并编辑。
- PowerPoint 直接插入原始 SVG，在幻灯片边界内按源比例 `contain` 居中，不改写 `viewBox`，不制造图内留白。

## DLL SHA256 与 ZIP 记录边界

以下哈希于 2026-07-29 在最终打包和标准本地安装后重新只读计算，三个 DLL 均在源码 Release、发布包和标准本地安装之间一致。

| 文件 | 源码 Release | 发布包 | 标准本地安装 |
| --- | --- | --- | --- |
| `DrawioPpt.Core.dll` | `AF5025768052D450D7B43E70229CA4FAEFECEBA50C3D1C33157F4FA7C560AED7` | `AF5025768052D450D7B43E70229CA4FAEFECEBA50C3D1C33157F4FA7C560AED7` | `AF5025768052D450D7B43E70229CA4FAEFECEBA50C3D1C33157F4FA7C560AED7` |
| `DrawioPpt.PowerPointAddIn.dll` | `4717622B7F8995F437989E3D9C2FF14A33C918EA9848A4702ADAFA095753931F` | `4717622B7F8995F437989E3D9C2FF14A33C918EA9848A4702ADAFA095753931F` | `4717622B7F8995F437989E3D9C2FF14A33C918EA9848A4702ADAFA095753931F` |
| `DrawioPpt.WordAddIn.dll` | `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8` | `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8` | `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8` |

相对位置：

- 源码 Release：`src\<project>\bin\x64\Release\<dll>`
- 发布包：`artifacts\releases\v1.0.8\package\bin\<dll>`
- 标准本地安装：`%LOCALAPPDATA%\Greensoft\DrawioPpt\bin\<dll>`

最终发布包已经包含 `A6AB…492E8` Word DLL，并已覆盖安装到标准本地目录；安装后的 Word、预览提供器和 PowerPoint 比例 E2E 再次通过。

ZIP SHA256 不写入会再次进入同一 ZIP 的本证据文档，避免自引用导致哈希必然变化；由发布脚本或包外交付摘要在最终打包完成后记录。

## Word UI 线程压力证据

- 首轮修正后结果：`artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json`，`ReportVersion=9`，SHA256 `0C4B1FFB8C2ABB31505B277A1035E2716D87E7456BF47C7030DBFD91618A5B1B`。
- 确认复跑结果：`artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-rerun.json`，`ReportVersion=9`，SHA256 `0D153F545245F0F0D6E405AA717E22012FE6916724722D14B6E111B6254F7D16`。
- 修正不透明位图检测后的最新结果：`artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-final.json`，`ReportVersion=9`，SHA256 `4C7B7C5EDA65B7F5BBF0A56FE18D974C7AA863B1C84D7E3556DF58D355A90C42`。
- 测试图片：1240×1753 PNG；独立 SVG 源比例 `0.707123`，PNG 比例 `0.707359`，误差 `0.000236`；`Front` 浮动布局。PNG 是 24bpp 不透明位图，透明留白检测明确记录 `BlankPaddingDetectionSupported=false` 和空的 `BorderPaddingPassed`，不冒充像素检测通过；导出命令使用 `--border 0`。普通图与受管图尺寸、锚点和环绕语义一致。
- 最新轮普通/受管选择 P95：`68.866/74.243 ms`；位置 P95：`321.844/325.764 ms`；缩放 P95：`89.633/122.968 ms`；目标选区事件 `96/96`，缺失 0。
- 最新轮 `ComparisonPassed=false`、`NoAdditionalMetadataPathDifferenceObserved=false`；受管图两轮位置 P95 为 `303.892/337.339 ms`，普通对照图为 `328.256/318.748 ms`。此前确认轮曾通过相对门槛，但不能用来覆盖最新失败。
- `AbsoluteLatencyGatePassed=false`、`OverallPassed=false`；最新轮最终失败和 Word 无响应采样均为 0。
- `PointerInputCovered=false`；自动化没有注入真实鼠标输入。本证据不声称鼠标拖动通过；用户已决定略过该人工验收，状态不是“通过”。
- 发布用[脱敏机器可读摘要](./evidence/word-ui-thread-stress-v1.0.8.json)已移除用户名、绝对附件路径、临时 GUID 和进程身份。

## 报告与日志

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-rerun.json`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-final.json`
- [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md)
- [Word UI 线程压力测试报告](./word-ui-thread-stress-report-v1.0.8.md)
- [Word UI 线程压力脱敏证据](./evidence/word-ui-thread-stress-v1.0.8.json)

## 已知边界

最终包、标准本地安装和最新源码 DLL 已同步，功能性 `12/12 PASS` 已覆盖本次 Word 增强。最新独立压力测试的受管图和普通对照图都存在超过 300ms 的轮次，且相对门槛失败，`OverallPassed=false`；真实指针验收由用户决定略过，不得写成通过。最终 ZIP 哈希在包外交付摘要记录。
