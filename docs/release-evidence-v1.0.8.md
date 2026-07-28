# v1.0.8 发布证据与日志索引

[中文](./release-evidence-v1.0.8.md) | [English](./release-evidence-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 交付物

- 版本：`v1.0.8`
- 发布目录：`artifacts\releases\v1.0.8\package`
- 发布压缩包：`artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip`
- 标准本地安装根：`%LOCALAPPDATA%\Greensoft\DrawioPpt`
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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" -DurationSeconds 12 -MaximumP95Ratio 1.25 -MaximumPerRoundP95Ratio 1.50 -MaximumSelectionP95DeltaMs 100 -MaximumAbsoluteSelectionP95Ms 2000 -MaximumAbsoluteP95Ms 2000 -MaximumAbsoluteResizeP95Ms 750 -MinimumOperationsPerRound 10 -WarmupSeconds 4 -ResizeOperationsPerRound 12 -SelectionSettleMilliseconds 75 -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-v1.0.8.json"
```

## 最终验证结果

- `Release|x64` 构建通过：0 warnings，0 errors。
- 所有源码和真实 Word 发布前门槛通过；详细关键值见 [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md)。
- 临时安装发布包的完整 Office E2E 为 9/9 通过：`ReleasePackage`、`InstallRelease`、`InstalledOfficeAddInsVerify`、`PowerPointAddInLoad`、`InstalledUrlSmoke`、`InstalledWordUrlHostE2E`、`InstalledWordComplexMetadataE2E`、`PowerPointUrlE2E`、`DesktopExporter` 全部为 PASS。
- 发布包包含 `DrawioPpt.WordAddIn.dll` 和 `scripts\word-complex-metadata-e2e.ps1`，并显式包含中英文架构、回归清单、优化 backlog、v1.0.5 历史 E2E、本版本发布说明、E2E 报告和证据索引；`PACKAGE.txt` 同时列出 README 与清单本身。
- 发布前压缩阶段的包内链接门槛输出 `PackageMarkdownMissingLinkCount=0`。
- full E2E 按 PID 和启动时间精确识别测试拥有的 Word；主脚本 AST 只有尾部一个 `exit 1`，构建、打包、安装和编译失败全部抛出后进入统一 catch/finally。受控 probe 验证主错误与卸载、仓库重注册、设置恢复三项错误会完整聚合到 FatalError 行，最终统一返回 1。
- 最终完整 E2E 的卸载临时包、还原仓库 Release 加载项注册和设置均成功，输出 `FullE2ECleanupSucceeded=True`，且没有 `WINWORD` / `POWERPNT` 残留。
- 上一条“还原仓库注册”是临时安装 E2E 结束时的历史状态；随后执行的标准本地安装已把 Word / PowerPoint `CodeBase` 切换到 `%LOCALAPPDATA%\Greensoft\DrawioPpt\bin`，`LoadBehavior=3`，并通过真实 COM 加载检查。
- 标准安装的三项 DLL 与发布包 SHA256 一致；复杂元数据、零被动选区路径与显式识别、实际 Word URL 宿主安装态回归均通过。真实宿主输出 `ExplicitEditAutomationAvailable=True`、`ExplicitEditCommandInvoked=True`，确认测试走的是“选中后显式重新编辑”，不是选区自动打开。
- 安装态 Word UI 线程压力对照执行两轮交叉顺序测试，每张图片每轮预热 4 秒后测量 12 秒，并分别记录选择、位置和尺寸延迟。普通/受管合并 P95 为 7.501 / 5.833 ms、388.910 / 368.568 ms、112.280 / 103.247 ms，受管/普通比例为 0.778 / 0.948 / 0.920。比较、绝对延迟、最小样本、最终失败和无响应门槛均通过；测量实例加载项已连接，按目标 Shape 分类确认 135 次事件且缺失为 0，安装态运行时反射输出 `NoPassiveSelectionMetadataPath=True`；保存重开后轻量引用为 382 字符，主存储恢复 254,606 字符源 XML。本样本不证明绝对延迟来源。

## 发布包 DLL SHA256

| 文件 | SHA256 |
| --- | --- |
| `bin\DrawioPpt.Core.dll` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` |
| `bin\DrawioPpt.PowerPointAddIn.dll` | `09E4BBE61B905BACA7B924868696A724A59EF37D3417F58E580AB7A01894733A` |
| `bin\DrawioPpt.WordAddIn.dll` | `6587C12F23709E5102DE5DD13DBD647CDA45CBDFB778E5D82FA12FA70761219D` |

压缩包 SHA256 不写入会被再次打入压缩包的文档，避免自哈希循环；它在最终重打包后单独记录。

## 报告与日志

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\logs\v1.0.8\drawioppt-full-e2e.log`
- [v1.0.8 Word UI 线程压力验收报告](./word-ui-thread-stress-report-v1.0.8.md)
- [Word UI 线程压力路径规范化证据快照](./evidence/word-ui-thread-stress-v1.0.8.json)

## 已知测试前提与边界

真实 Office 测试开始前 Word 和 PowerPoint 必须关闭，且不能与其他 Office 自动化并行。标准目录本地安装和真实 Word UI 线程压力对照已完成；当前交互桌面因 Win32 error 5 无法注入指针动作，所以普通/受管图形各连续 10 秒的人工指针拖动、缩放、重定位和“重新编辑”按钮点击仍待用户确认。本证据不把自动化压力结果表述为人工鼠标验收。
