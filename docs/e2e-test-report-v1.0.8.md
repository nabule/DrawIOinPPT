# v1.0.8 E2E 测试报告

[中文](./e2e-test-report-v1.0.8.md) | [English](./e2e-test-report-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 测试环境

- 日期：2026-07-28
- Windows + Microsoft Word Desktop / PowerPoint Desktop
- .NET Framework 4.8，`Release|x64`
- 当前用户级 Office COM Add-in 注册和临时安装的 v1.0.8 发布包

## 发布前门槛

以下命令按顺序串行执行，开始前确认没有用户 Word / PowerPoint 进程，结束后确认没有 `WINWORD` / `POWERPNT` 残留。

| 项目 | 结果 | 关键证据 |
| --- | --- | --- |
| Release x64 构建 | 通过 | 0 warnings，0 errors |
| Word 零被动选区路径源码约束 | 通过 | `WORD_SELECTION_SYNC_TEST_PASS`：无 `WindowSelectionChange` / 自动打开 / 启动时选区读取；四个显式命令一次捕获同一实时图片 |
| WebView2 用户目录源码约束 | 通过 | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Word URL 宿主测试安全约束 | 通过 | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| full E2E 清理与统一出口安全约束 | 通过 | `FULL_E2E_CLEANUP_SAFETY_TEST_PASS`；主脚本 AST 只有最终一个 `exit 1`，错误启动时间不会终止进程，受控 probe 将主错误及卸载、重注册、设置恢复错误全部写入 FatalError 行并返回 1 |
| Word 复杂元数据 E2E | 通过 | `DrawioXmlChars=271361`、`AlternativeTextChars=436`，主存储、轻量引用、失败回退、旧版迁移、part ID 稳定、保存关闭重开均满足断言 |
| Word 显式选区识别 E2E | 通过 | `NoPassiveSelectionMetadataPath=True`、`ExplicitManagedSelectionDetected=True`、`ExplicitPlainPictureCanBind=True` |
| Word URL E2E | 通过 | `CreatedManagedPicture=True`、`ReopenedEditApplied=True`、`PersistedAfterReopen=True` |
| 真实 Word URL 宿主 E2E | 通过 | `WordAddInConnect=True`、`ActualHostProcess=WINWORD`、`ExplicitEditAutomationAvailable=True`、`ExplicitEditCommandInvoked=True`、`ActualWordUrlEditorSaved=True`、`AccessDeniedDialog=False` |

复杂元数据回归还验证了 `StoredPayload=True`、`LightweightAlternativeText=True`、`FallbackRetainsPayload=True`、`LegacyMigrationPassed=True`、`LegacyMigrationPartIdsStable=True`、`FallbackPersistedAfterReopen=True`、`LegacyMigrationPersistedAfterReopen=True`、`CleanupResidualWinWord=False`。

## 临时安装包完整 Office E2E

执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

结果：9/9 通过。

打包阶段先输出 `PackageMarkdownMissingLinkCount=0`，确认 ZIP 内所有本地 Markdown 链接均能解析到包内文件。

| 检查项 | 结果 | 关键证据 |
| --- | --- | --- |
| `ReleasePackage` | PASS | v1.0.8 发布目录生成成功 |
| `InstallRelease` | PASS | 发布包安装到隔离临时目录 |
| `InstalledOfficeAddInsVerify` | PASS | Word / PowerPoint 注册和 COM 加载验证通过 |
| `PowerPointAddInLoad` | PASS | `Connect=True` |
| `InstalledUrlSmoke` | PASS | `Saved=True`、`SvgHasSmoke=True`、`XmlHasSmokeId=True` |
| `InstalledWordUrlHostE2E` | PASS | 临时安装包 DLL 在真实 `WINWORD.EXE` 中通过选中后显式命令完成 URL 回写 |
| `InstalledWordComplexMetadataE2E` | PASS | 临时安装包 DLL 完成复杂 XML 主存储、轻量引用、回退和迁移回归 |
| `PowerPointUrlE2E` | PASS | `CreatedManagedShape=True`、`ReopenedEditApplied=True`、`PersistedAfterReopen=True`、`ExitCode=0` |
| `DesktopExporter` | PASS | draw.io Desktop 导出 SVG 成功 |

URL smoke 的 mock 编辑器会在 preferred loopback 端口被系统排除或占用时连续尝试后续端口；探活和 `PluginSettings.EditorUrl` 都使用最终绑定端口。本次 smoke 输出 `MockServerPort=8752`，真实 Word URL 宿主输出 `MockServerPort=12440`。

安装态复杂元数据测试写出 `TestWordProcessIdentity=42848:639208283590231895`。finally 按 PID 和启动时间核对测试拥有的 Word，本次进程已正常退出，因此输出 `TestOwnedProcessAlreadyExited=42848`；卸载临时包、恢复仓库加载项注册和设置均完成，最终输出 `FullE2ECleanupSucceeded=True`。

## 后续标准本地安装与 Word UI 线程压力确认

- 标准安装根：`%LOCALAPPDATA%\Greensoft\DrawioPpt`，`PACKAGE.txt` 为 `v1.0.8 / Release / x64`。
- 标准安装目录的 Core、PowerPoint 和 Word DLL 与发布包 SHA256 一致；Word / PowerPoint `CodeBase` 均指向标准安装目录，`LoadBehavior=3`。
- `verify-office-install.ps1`、`word-addin-load-check.ps1`、已安装版 `word-complex-metadata-e2e.ps1`、`word-selection-event-e2e.ps1` 和实际 Word URL 宿主回归分别通过；这些是标准安装确认，不计入上面的临时安装 9/9。
- 使用同一 120 元素复杂 SVG 创建相同尺寸的普通图形与受管图形。受管图形引用为 382 字符，完整 Draw.io XML 254,606 字符保存在 `Document.CustomXMLParts`。
- 在可见 Word 中执行两轮交叉顺序测试，每张图片每轮先预热 4 秒，再测量 12 秒；选择事件和位置更新分别计时。普通/受管合并选择 P95 为 7.501 / 5.833 ms，位置 P95 为 388.910 / 368.568 ms，尺寸 P95 为 112.280 / 103.247 ms；受管/普通比例分别为 0.778 / 0.948 / 0.920。比较、绝对延迟、最小样本、最终调用失败和无响应采样门槛均通过；测量实例加载项已连接，测试计数器按目标 Shape 分类确认 135 次事件且缺失 0，安装态运行时反射输出 `NoPassiveSelectionMetadataPath=True`。
- 保存并重开后，受管图形位置/尺寸为 `69 / 82 / 340 / 226`，轻量引用仍为 382 字符，文档主存储恢复 254,606 字符 XML，`WordAddInConnect=True`。
- 本样本只支持“未观察到图片元数据路径带来超过门槛的额外差异，且绝对延迟未超过本次验收门槛”，不证明绝对延迟的具体来源。完整命令、方法和路径规范化证据快照见 [Word UI 线程压力验收报告](./word-ui-thread-stress-report-v1.0.8.md)。

## 报告与日志

- 汇总报告：`artifacts\test-reports\full-e2e-v1.0.8.md`
- 报告副本：`artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- PowerShell transcript：`artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- 插件日志快照：`artifacts\logs\v1.0.8\drawioppt-full-e2e.log`
- 标准安装 UI 线程压力报告：[v1.0.8 Word UI 线程压力验收报告](./word-ui-thread-stress-report-v1.0.8.md)
- 机器可读压力结果：[路径规范化证据快照](./evidence/word-ui-thread-stress-v1.0.8.json)
- 发布内容与 DLL 哈希：见 [v1.0.8 发布证据](./release-evidence-v1.0.8.md)

## 验收边界

本报告验证源码 Release DLL、临时安装发布包、标准本地安装及真实 Word UI 线程压力对照。当前 Codex 交互桌面的光标 API 被系统以 Win32 error 5 拒绝，无法执行真实指针连续拖动；普通/受管图形各连续 10 秒的人工指针拖动、缩放、重定位和“重新编辑”按钮点击仍是独立待确认项，不计入自动化 9/9 或 UI 线程压力结果。
