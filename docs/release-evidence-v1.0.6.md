# v1.0.6 发布证据与日志索引

[中文](./release-evidence-v1.0.6.md) | [English](./release-evidence-v1.0.6.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 交付物

- 版本：`v1.0.6`
- 发布目录：`artifacts\releases\v1.0.6\package`
- 发布压缩包：`artifacts\releases\v1.0.6\DrawioPpt-v1.0.6.zip`
- 新增发布验证脚本：`scripts\word-selection-event-e2e.ps1`

## 发布门槛

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.6 -Configuration Release -Platform x64 -SkipBuild
```

发布包必须包含 `DrawioPpt.WordAddIn.dll`、`word-selection-event-e2e.ps1`、本版本中英文发布说明、E2E 报告和本证据索引。发布 DLL 使用以下命令验证：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.6\package\scripts\word-selection-event-e2e.ps1 -SkipBuild -SkipSourceCheck -AssemblyRoot .\artifacts\releases\v1.0.6\package\bin
```

随后对临时安装目录运行 `verify-office-install.ps1`、`word-addin-load-check.ps1` 和完整 Office E2E。命令输出与 GitHub Release 资产 SHA256 作为本次发布记录的一部分。

## 最终验证结果

- `Release|x64` 构建通过：0 warnings，0 errors。
- 源码与 Release DLL 的 `word-selection-event-e2e.ps1` 均通过，输出 `NoSelectionPolling=True`、`ManagedSelectionDetected=True`、`PlainPictureCanBind=True`。
- `word-url-e2e.ps1` 通过：mock HTTP 可达、URL 图新建、重开编辑和再次重开持久化均为 `True`；加载项设置恢复为 `LoadBehavior=3`。
- 发布包已安装到临时目录；`verify-office-install.ps1` 验证 Word/PowerPoint 注册和实际 COM 加载，`word-addin-load-check.ps1` 输出 `WordAddInConnect=True`。
- 完整 Office E2E 报告为 7/7 通过：发布包、安装、Office 验收、PowerPoint 加载、已安装 URL smoke、PowerPoint URL E2E、桌面 draw.io 导出；Office 进程清理检查通过。

## 已知测试前提

Word COM E2E 要求 Word 已关闭，且不能与其他 Word 自动化并行执行。`word-url-e2e.ps1` 会在测试期间临时设置当前用户 Word 加载项的 `LoadBehavior=0` 以隔离测试宿主，并在 finally 中恢复原值；运行期间不要启动 Word。
