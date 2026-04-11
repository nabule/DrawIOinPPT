# Full E2E Test Report (v1.0.0)

本报告对应 `DrawioPpt v1.0.0` 的实际发布验证结果，测试执行日期为 `2026-04-12`，目标是确认发布包、安装脚本、PowerPoint 宿主加载、URL 模式回写链路、重新打开后的继续编辑，以及桌面导出能力都能在当前机器上跑通。

## 测试环境

- 操作系统：`Microsoft Windows 10 专业版 10.0.19045 64-bit`
- PowerPoint：`Microsoft PowerPoint Desktop x64`
- .NET：`.NET Framework 4.8`
- WebView2 Runtime：已安装
- draw.io Desktop：`C:\Program Files\draw.io\draw.io.exe`
- 发布版本：`v1.0.0`
- 验证提交：`6239eab246aa603a90af119b640aaff1186f95dc`

## 执行过的实际命令

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1 -SkipBuild
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.0 -SkipBuild
```

## 实测覆盖范围

- 环境前提检查
- Release 构建
- URL 编辑器本地烟雾测试
- 发布包生成
- 发布包安装与当前用户注册
- PowerPoint COM Add-in 自动加载
- URL 模式新建图形并完成 `configure -> init -> load -> save -> export`
- 保存 PPT、关闭、重新打开后再次编辑同一图形
- `Shape.Tags / AlternativeText / Presentation.CustomXMLParts` 持久化验证
- draw.io Desktop CLI 导出 SVG 验证

## 关键观察

- URL 烟雾测试返回：
  - `DialogResult=OK`
  - `Saved=True`
  - `SvgHasSmoke=True`
  - `XmlHasSmokeId=True`
- 完整 E2E 中首次创建后：
  - `CreatedManagedShape=True`
  - `CreatedDiagramId=7a0c3b5b674d4934a99a74c6f7546543`
  - `CreatedCustomXmlCount=4`
- 关闭并重开后再次编辑：
  - `ReopenedEditApplied=True`
  - `EditedCustomXmlCount=4`
- 再次关闭并重开后的持久化检查：
  - `PersistedAfterReopen=True`
  - `FinalCustomXmlCount=4`
  - `ShapeCount=1`

这些结果说明 `v1.0.0` 当前机器上的 URL 模式回写、二次编辑和 `CustomXMLParts` 持久化链路均已通过真实 PowerPoint 宿主验证。

## 测试结果汇总

| Check | Result | Detail |
| --- | --- | --- |
| ReleasePackage | PASS | `artifacts\releases\v1.0.0\package` |
| InstallRelease | PASS | 发布包可安装到临时安装目录并完成注册 |
| PowerPointAddInLoad | PASS | `Connect=True` |
| InstalledUrlSmoke | PASS | 发布包内 `url-editor-smoke.ps1` 可运行通过 |
| PowerPointUrlE2E | PASS | `ExitCode=0` |
| DesktopExporter | PASS | `C:\Program Files\draw.io\draw.io.exe` |

## 产物与日志

- 原始 E2E 结果：`artifacts\test-reports\full-e2e-v1.0.0.md`
- 本版本日志目录：`artifacts\logs\v1.0.0\`
- 插件日志快照：`artifacts\logs\v1.0.0\drawioppt-full-e2e.log`
- E2E 控制台日志：`artifacts\logs\v1.0.0\full-e2e.console.log`
- URL 烟雾测试日志：`artifacts\logs\v1.0.0\url-editor-smoke.log`

## 结论

`v1.0.0` 在当前验证环境下已满足发布条件：程序可构建、发布包可生成和安装、PowerPoint 可正常加载插件，URL 模式的创建/回写/重开/再编辑链路已通过实际测试，桌面导出能力也已验证通过。
