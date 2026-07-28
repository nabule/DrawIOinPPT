# v1.0.8 发布证据与日志索引

[中文](./release-evidence-v1.0.8.md) | [English](./release-evidence-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 交付物

- 版本：`v1.0.8`
- 发布目录：`artifacts\releases\v1.0.8\package`
- 发布压缩包：`artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip`
- 新增已安装包回归：`scripts\word-complex-metadata-e2e.ps1`

v1.0.8 仍处于待公开发布状态；本仓库证据不表示 GitHub Release 已创建。

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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

## 最终验证结果

- `Release|x64` 构建通过：0 warnings，0 errors。
- 所有源码和真实 Word 发布前门槛通过；详细关键值见 [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md)。
- 临时安装发布包的完整 Office E2E 为 9/9 通过：`ReleasePackage`、`InstallRelease`、`InstalledOfficeAddInsVerify`、`PowerPointAddInLoad`、`InstalledUrlSmoke`、`InstalledWordUrlHostE2E`、`InstalledWordComplexMetadataE2E`、`PowerPointUrlE2E`、`DesktopExporter` 全部为 PASS。
- 发布包包含 `DrawioPpt.WordAddIn.dll` 和 `scripts\word-complex-metadata-e2e.ps1`，并显式包含中英文架构、回归清单、优化 backlog、v1.0.5 历史 E2E、本版本发布说明、E2E 报告和证据索引；`PACKAGE.txt` 同时列出 README 与清单本身。
- 发布前压缩阶段的包内链接门槛输出 `PackageMarkdownMissingLinkCount=0`。
- full E2E 按 PID 和启动时间精确识别测试拥有的 Word；主脚本 AST 只有尾部一个 `exit 1`，构建、打包、安装和编译失败全部抛出后进入统一 catch/finally。受控 probe 验证主错误与卸载、仓库重注册、设置恢复三项错误会完整聚合到 FatalError 行，最终统一返回 1。
- 第三次完整 E2E 的卸载临时包、还原仓库 Release 加载项注册和设置均成功，输出 `FullE2ECleanupSucceeded=True`，且没有 `WINWORD` / `POWERPNT` 残留。

## 发布包 DLL SHA256

| 文件 | SHA256 |
| --- | --- |
| `bin\DrawioPpt.Core.dll` | `C851D64AD79689EF83615E619C6A2CFD3E07599AD2EC888602A3F4136D8FADA6` |
| `bin\DrawioPpt.PowerPointAddIn.dll` | `C8425FC93C103C3CFAA80A0BF33468363ADE95165EE76308B98A5FA37F76786A` |
| `bin\DrawioPpt.WordAddIn.dll` | `4FE2ECE68A8F3762A5B930AB814DF64A41766D11FCA33DA12F8BBF3DF51A5BB8` |

压缩包 SHA256 不写入会被再次打入压缩包的文档，避免自哈希循环；它在最终重打包后单独记录。

## 报告与日志

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\logs\v1.0.8\drawioppt-full-e2e.log`

## 已知测试前提与边界

真实 Office 测试开始前 Word 和 PowerPoint 必须关闭，且不能与其他 Office 自动化并行。最终标准目录本地安装和人工拖动对比属于后续验收，本证据不提前宣称它们已经通过。
