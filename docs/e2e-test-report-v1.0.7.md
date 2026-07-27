# v1.0.7 E2E 测试报告

[中文](./e2e-test-report-v1.0.7.md) | [English](./e2e-test-report-v1.0.7.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 测试环境

- Windows + Microsoft Word Desktop / PowerPoint Desktop
- .NET Framework 4.8，`Release|x64`
- 当前用户级 Office COM Add-in 注册和已安装发布包

## 本次回归

| 项目 | 结果 | 关键证据 |
| --- | --- | --- |
| WebView2 用户目录源码约束 | 通过 | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Word URL 宿主测试安全约束 | 通过 | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| Release 构建 | 通过 | 0 warnings，0 errors |
| Word 选区事件 E2E | 通过 | `NoSelectionPolling=True`、`ManagedSelectionDetected=True`、`PlainPictureCanBind=True` |
| 真实 Word URL 宿主 E2E | 通过 | `WordAddInConnect=True`、`ActualHostProcess=WINWORD`、`ActualWordUrlEditorSaved=True`、`WebView2UserDataFolderExists=True` |
| 已安装发布 DLL 的完整 Office E2E | 通过（8/8） | 发布包、安装、Office 注册/加载、URL smoke、真实 Word URL 宿主、PowerPoint URL E2E、桌面导出均为 PASS |

## 覆盖说明

`word-url-addin-host-e2e.ps1` 不通过临时 EXE 宿主初始化 WebView2，而是在真实 `WINWORD.EXE` 已加载 Word COM Add-in 的前提下选择受管图片。脚本使用本地 mock URL 编辑器完成 `configure -> init -> load -> save -> export`，验证更新后的 XML 回写、`%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2` 存在，并检查本次日志没有 URL 初始化的 `E_ACCESSDENIED`。脚本会恢复 Word 加载项、COM 类注册和插件设置。

完整 Office E2E 报告位于 `artifacts\test-reports\full-e2e-v1.0.7.md`，其 8 项分别为 `ReleasePackage`、`InstallRelease`、`InstalledOfficeAddInsVerify`、`PowerPointAddInLoad`、`InstalledUrlSmoke`、`InstalledWordUrlHostE2E`、`PowerPointUrlE2E` 与 `DesktopExporter`，全部通过。发布包内容与发布资产校验见 [发布证据与日志索引](./release-evidence-v1.0.7.md)。
