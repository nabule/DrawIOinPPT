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
| Word 选区同步源码约束 | 通过 | `WORD_SELECTION_SYNC_TEST_PASS` |
| WebView2 用户目录源码约束 | 通过 | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Word URL 宿主测试安全约束 | 通过 | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| full E2E 清理与统一出口安全约束 | 通过 | `FULL_E2E_CLEANUP_SAFETY_TEST_PASS`；主脚本 AST 只有最终一个 `exit 1`，错误启动时间不会终止进程，受控 probe 将主错误及卸载、重注册、设置恢复错误全部写入 FatalError 行并返回 1 |
| Word 复杂元数据 E2E | 通过 | `DrawioXmlChars=271361`、`AlternativeTextChars=436`，主存储、轻量引用、失败回退、旧版迁移、part ID 稳定、保存关闭重开均满足断言 |
| Word 选区事件 E2E | 通过 | `NoSelectionPolling=True`、`ManagedSelectionDetected=True`、`PlainPictureCanBind=True` |
| Word URL E2E | 通过 | `CreatedManagedPicture=True`、`ReopenedEditApplied=True`、`PersistedAfterReopen=True` |
| 真实 Word URL 宿主 E2E | 通过 | `WordAddInConnect=True`、`ActualHostProcess=WINWORD`、`ActualWordUrlEditorSaved=True`、`WebView2UserDataFolderExists=True`、`AccessDeniedDialog=False` |

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
| `InstalledWordUrlHostE2E` | PASS | 临时安装包 DLL 在真实 `WINWORD.EXE` 中完成 URL 回写 |
| `InstalledWordComplexMetadataE2E` | PASS | 临时安装包 DLL 完成复杂 XML 主存储、轻量引用、回退和迁移回归 |
| `PowerPointUrlE2E` | PASS | `CreatedManagedShape=True`、`ReopenedEditApplied=True`、`PersistedAfterReopen=True`、`ExitCode=0` |
| `DesktopExporter` | PASS | draw.io Desktop 导出 SVG 成功 |

PowerPoint mock 编辑器先由操作系统分配 preferred loopback 端口，并在 C# 宿主中使用 `StartWithRetry` 处理绑定竞争；探活和 `PluginSettings.EditorUrl` 都使用最终绑定端口。本次输出为 `MockServerPort=6508`。

安装态复杂元数据测试写出 `TestWordProcessIdentity=23584:639208157549482875`。finally 按 PID 和启动时间核对测试拥有的 Word，本次进程已正常退出，因此输出 `TestOwnedProcessAlreadyExited=23584`；卸载临时包、恢复仓库加载项注册和设置均完成，最终输出 `FullE2ECleanupSucceeded=True`。

## 报告与日志

- 汇总报告：`artifacts\test-reports\full-e2e-v1.0.8.md`
- 报告副本：`artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- PowerShell transcript：`artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- 插件日志快照：`artifacts\logs\v1.0.8\drawioppt-full-e2e.log`
- 发布内容与 DLL 哈希：见 [v1.0.8 发布证据](./release-evidence-v1.0.8.md)

## 验收边界

本报告验证的是源码 Release DLL 和临时安装的发布包。标准本地目录的最终安装与人工 10 秒拖动、缩放、重定位和重开编辑对比属于后续验收，不在本报告中提前宣称通过。
