# v1.0.8 E2E 测试报告

[中文](./e2e-test-report-v1.0.8.md) | [English](./e2e-test-report-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 测试环境

- 日期：2026-07-29
- Windows + Microsoft Word Desktop / PowerPoint Desktop
- .NET Framework 4.8，`Release|x64`
- 当前用户级 Office COM Add-in 注册、隔离临时安装的 v1.0.8 发布包和标准本地安装

## 发布前门槛

真实 Office 命令串行执行；开始前要求没有用户 Word / PowerPoint 进程，结束后要求没有 `WINWORD` / `POWERPNT` 残留。

| 项目 | 结果 | 关键证据 |
| --- | --- | --- |
| Release x64 构建 | 通过 | 0 warnings，0 errors |
| Word 零被动选区路径源码约束 | 通过 | `WORD_SELECTION_SYNC_TEST_PASS`：无 `WindowSelectionChange`、自动打开或启动时选区读取；四个显式命令各捕获一次实时图片 |
| WebView2 用户目录源码约束 | 通过 | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Word URL 宿主测试安全约束 | 通过 | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| full E2E 清理安全约束 | 通过 | `FULL_E2E_CLEANUP_SAFETY_TEST_PASS`：安装前快照 Word / PowerPoint 注册树；finally 精确恢复值、类型、子键和原本不存在状态；最终残留 Office PID 会写入 FatalError 并令 E2E 失败 |
| Word 复杂元数据 E2E | 通过 | 文档级主存储、轻量引用、失败回退、旧版迁移、稳定 part ID 和保存关闭重开均满足断言 |
| Word 显式选区识别 E2E | 通过 | `NoPassiveSelectionMetadataPath=True`、`ExplicitManagedSelectionDetected=True`、`ExplicitPlainPictureCanBind=True` |
| 真实 Word URL 宿主 E2E | 通过 | `WordAddInConnect=True`、`ExplicitEditAutomationAvailable=True`、`ExplicitEditCommandInvoked=True`、`ActualWordUrlEditorSaved=True` |

显式命令自动化验证的是“选中图片后调用重新编辑命令”，不是选中自动打开，也不等同于用户真实鼠标点击验收。

## 临时安装包完整 Office E2E

执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

结果：`12/12 PASS`。

| 检查项 | 结果 | 关键证据 |
| --- | --- | --- |
| `ReleasePackage` | PASS | v1.0.8 发布目录生成成功 |
| `InstallRelease` | PASS | 发布包安装到隔离临时目录 |
| `InstalledOfficeAddInsVerify` | PASS | Word / PowerPoint 注册和 COM 加载通过 |
| `PowerPointAddInLoad` | PASS | `Connect=True` |
| `InstalledUrlSmoke` | PASS | URL 编辑器 smoke 完成保存和 SVG/XML 回写 |
| `InstalledWordUrlHostE2E` | PASS | 临时包 DLL 在真实 `WINWORD.EXE` 中通过选中后显式命令完成 URL 回写 |
| `InstalledWordComplexMetadataE2E` | PASS | 复杂 XML 主存储、轻量引用、回退和迁移回归通过 |
| `InstalledWordSvgAspectE2E` | PASS | 包快照中的 Word 2:1 图片新建、替换和旧图比例修复通过，保持浮动 `wdWrapFront` |
| `InstalledWordPreviewProviderE2E` | PASS | 包快照中的 Word 620px 等比例 PNG、零方形强制、临时源清理和安全回退通过 |
| `InstalledPowerPointSvgAspectE2E` | PASS | PowerPoint 原始 SVG `contain` 等比居中且不扩展 `viewBox` |
| `PowerPointUrlE2E` | PASS | PowerPoint URL 创建、重开、编辑和持久化通过 |
| `DesktopExporter` | PASS | draw.io Desktop 导出 SVG 成功 |

打包阶段输出 `PackageMarkdownMissingLinkCount=0`。汇总报告和日志副本内容一致，SHA256 为 `7D5AC38245C29AF08B04C24414480E1024E270210657EC0AFB97D0ABF52CF73C`。

临时安装结束后卸载测试包，并把 Word / PowerPoint 加载项注册精确恢复到执行前状态，包括执行前不存在的键。设置恢复完成，最终输出 `FullE2ECleanupSucceeded=True`，没有 `WINWORD` / `POWERPNT` 残留。

最终包已覆盖增强后的真实 Word 回归：`VerticalRatio=0.25`、`VerticalContained=True`、`FailedReplacementPreservedOriginal=True`。预览门禁覆盖 draw.io Desktop 导出失败时的 620×310 等比例轻量 PNG 占位预览，以及临时 XML 有限重试、后台延迟清理和启动清理；不再回退复杂 SVG。

## 标准本地安装与制品一致性

- 标准安装根：`%LOCALAPPDATA%\Greensoft\DrawioPpt`；`PACKAGE.txt` 为 `v1.0.8 / Release / x64`。
- Core、PowerPoint 与 Word DLL 在源码 Release 输出、发布包和标准本地安装中 SHA256 一致；Word DLL 为 `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D`。
- PowerPoint 直接插入原始 SVG，按幻灯片边界 `contain` 等比居中，不改写 `viewBox`、不制造图内留白。
- 最新源码中 Word 正常展示为 620px 等比例 PNG，插入为 `wdWrapFront` 浮动 `Shape`；横向/纵向图均 `contain`，替换先建新图后删原图，失败时保留原图。导出失败则生成 620×310 等比例轻量 PNG 占位图，源 XML 仍在文档元数据中。选中本身不编辑，点击显式命令后才读取选区并进入编辑。
- ZIP SHA256 不写入会再次进入压缩包的本报告，避免自引用失真；由发布脚本或包外交付摘要在最终重打包后记录。

## Word UI 线程压力结果

最新用户复杂源图压力运行使用 620×876 PNG、`Front` 浮动布局和富文档基线。普通/受管位置 P95 为 `524.962/488.881 ms`，受管图低 `36.081 ms`；目标 Shape 选区事件 `98/98`，缺失 0。保存重开、几何、存储和清理门槛通过。

绝对位置 P95 门槛为 `300 ms`，普通图和受管图均失败；`17×24 px` 纯色 PNG 基线也失败。机器结果为 `SelectionComparisonPassed=false`、`ComparisonPassed=false`、`AbsoluteLatencyGatePassed=false`、`NoAdditionalMetadataPathDifferenceObserved=false`、`OverallPassed=false`。因此不能把绝对延迟归因于图形复杂度或受管元数据，也不能写成性能验收通过。

`PointerInputCovered=false`。Windows `GetCursorPos` 访问被拒绝，`SetIsBorderRequired` 接口不受支持；自动化没有覆盖真实鼠标输入，不作鼠标拖动通过结论。

## 报告与日志

- 功能 E2E 汇总：`artifacts\test-reports\full-e2e-v1.0.8.md`
- 日志副本：`artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- PowerShell transcript：`artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- Word 压力原始结果：`artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json`
- [Word UI 线程压力测试报告](./word-ui-thread-stress-report-v1.0.8.md)
- [脱敏机器可读压力证据](./evidence/word-ui-thread-stress-v1.0.8.json)
- [发布内容与哈希](./release-evidence-v1.0.8.md)

## 验收边界

本报告记录的最终包功能性 Office E2E 为 `12/12 PASS`，标准本地安装也已更新并复核。Word 压力门槛为独立结果，当前 `OverallPassed=false`。真实指针拖动、缩放、重定位和人工点击“重新编辑”仍需用户确认。
