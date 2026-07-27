# v1.0.7 发布证据与日志索引

[中文](./release-evidence-v1.0.7.md) | [English](./release-evidence-v1.0.7.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 交付物

- 版本：`v1.0.7`
- 发布目录：`artifacts\releases\v1.0.7\package`
- 发布压缩包：`artifacts\releases\v1.0.7\DrawioPpt-v1.0.7.zip`
- 新增发布验证脚本：`scripts\word-url-addin-host-e2e.ps1`

## 发布门槛

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\webview2-user-data-folder-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.7
```

## 最终验证结果

- `Release|x64` 构建通过：0 warnings，0 errors。
- `webview2-user-data-folder-test.ps1` 通过，输出 `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS`。
- `word-url-addin-host-e2e-safety-test.ps1` 通过，输出 `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS`；它约束 mock 成功绑定后才写入 URL 设置、绑定失败时重试可用 loopback 端口、按 helper 即时落盘的 PID/启动时间清理，以及互相隔离的恢复步骤。
- 源码 Release DLL 的 `word-url-addin-host-e2e.ps1` 通过：真实 `WINWORD.EXE` 中 `WordAddInConnect=True`、`ActualWordUrlEditorSaved=True`、`WebView2UserDataFolderExists=True`，且本次日志无 URL 初始化 `E_ACCESSDENIED`。
- `word-selection-event-e2e.ps1` 通过：`NoSelectionPolling=True`、`ManagedSelectionDetected=True`、`PlainPictureCanBind=True`。
- 完整 Office E2E 为 8/8 通过，其中 `InstalledWordUrlHostE2E` 从临时安装发布包的 `bin` 注册 DLL 并在真实 Word 宿主中完成 URL 回写。
- 最终发布包必须包含 `DrawioPpt.WordAddIn.dll`、`word-url-addin-host-e2e.ps1`、两项源码约束脚本、本版本中英文发布说明、E2E 报告与本证据索引。压缩包 SHA256 作为 GitHub Release 资产校验记录。

## 已知测试前提

真实 Word URL 宿主测试开始前 Word 必须关闭，且不要与其他 Word 自动化并行运行。脚本只终止自身创建的 Word 进程，并在 finally 中还原 Word 加载项、COM 类注册和插件设置。
