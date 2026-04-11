# v1.0.0 发布证据与日志索引

本文档汇总 `DrawioPpt v1.0.0` 的程序交付物、实际测试结果和日志位置，便于发布、验收和问题追溯。

## 发布信息

- 版本：`v1.0.0`
- 发布日期：`2026-04-12`
- 发布提交：`6239eab246aa603a90af119b640aaff1186f95dc`
- 程序集版本：
  - `DrawioPpt.Core = 1.0.0.0`
  - `DrawioPpt.PowerPointAddIn = 1.0.0.0`

## 程序交付物

- 发布目录：`artifacts\releases\v1.0.0\package`
- 发布压缩包：`artifacts\releases\v1.0.0\DrawioPpt-v1.0.0.zip`
- 主要二进制：
  - `bin\DrawioPpt.PowerPointAddIn.dll`
  - `bin\DrawioPpt.Core.dll`
  - `bin\Microsoft.Web.WebView2.Core.dll`
  - `bin\Microsoft.Web.WebView2.WinForms.dll`
  - `bin\runtimes\win-x64\native\WebView2Loader.dll`
- 安装入口：
  - `install.cmd`
  - `scripts\install-release.ps1`
- 卸载入口：
  - `uninstall.cmd`
  - `scripts\uninstall-release.ps1`

## 文档交付物

- [安装与联调说明](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/installation.md)
- [用户手册](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/user-guide.md)
- [回归清单](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/regression-checklist.md)
- [v1.0.0 发布说明](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/release-notes-v1.0.0.md)
- [v1.0.0 E2E 测试报告](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/e2e-test-report-v1.0.0.md)

## 实际测试情况

测试执行日期：`2026-04-12`

测试环境：

- 操作系统：`Microsoft Windows 10 专业版 10.0.19045 64-bit`
- PowerPoint：`Microsoft PowerPoint Desktop x64`
- draw.io Desktop：`C:\Program Files\draw.io\draw.io.exe`
- WebView2 Runtime：已安装

已实际执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1 -SkipBuild
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.0 -SkipBuild
```

测试结论：

- 环境检查通过
- Release 构建通过，`0` 警告、`0` 错误
- URL 烟雾测试通过
- 发布包生成、安装、注册和 PowerPoint 宿主加载通过
- URL 模式创建、保存、关闭重开、再次编辑、再次重开后的持久化验证通过
- draw.io Desktop CLI 导出能力通过

关键结果摘录：

- `CreatedManagedShape=True`
- `ReopenedEditApplied=True`
- `PersistedAfterReopen=True`
- `FinalCustomXmlCount=4`
- `ShapeCount=1`

## 日志索引

本版本全部发布相关日志保存在：

```text
artifacts\logs\v1.0.0\
```

关键日志文件：

- `dev-check.log`
  环境检查输出
- `build-release.log`
  Release 构建输出
- `url-editor-smoke.log`
  URL 模式烟雾测试控制台输出
- `drawioppt-url-editor-smoke.log`
  URL 烟雾测试阶段插件日志快照
- `full-e2e.console.log`
  完整 E2E 控制台输出
- `full-e2e-v1.0.0.transcript.log`
  完整 E2E PowerShell transcript
- `full-e2e-v1.0.0.report.md`
  原始 E2E 报告副本
- `drawioppt-full-e2e.log`
  完整 E2E 期间插件日志快照

## 交付建议

- 对外分发时优先提供 `DrawioPpt-v1.0.0.zip`
- 若需要验收依据，连同 `docs\release-evidence-v1.0.0.md`、`docs\e2e-test-report-v1.0.0.md` 和 `logs\` 目录一起提供
- 若现场排障，优先查看 `logs\drawioppt-full-e2e.log` 和用户本机 `%APPDATA%\Greensoft\DrawioPpt\Logs\drawioppt.log`
