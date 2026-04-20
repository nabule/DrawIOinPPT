# Full E2E Test Report (v1.0.1)

[中文](./e2e-test-report-v1.0.1.md) | [English](./e2e-test-report-v1.0.1.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本报告记录 `DrawioPpt v1.0.1` 的实际发布验证结果。重点包括真实发布链路验证，以及本版本最关键的 SVG 矢量文字回归验证。

## 测试环境

- 操作系统：`Microsoft Windows 10 专业版 10.0.19045 64-bit`
- PowerPoint：`Microsoft PowerPoint Desktop x64`
- .NET：`.NET Framework 4.8`
- WebView2 Runtime：已安装
- draw.io Desktop：`C:\Program Files\draw.io\draw.io.exe`
- 发布版本：`v1.0.1`

## 执行过的实际命令

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\restore-packages.ps1
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe .\DrawioPpt.sln /t:Rebuild /p:Configuration=Release /p:Platform=x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.1 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.1\package\scripts\install-release.ps1 -InstallRoot C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1\scripts\url-editor-smoke.ps1 -SkipBuild
```

以上步骤是本次版本的真实发布验收路径，覆盖了构建、打包、安装、URL 编辑器宿主联通和 SVG 文字矢量回归。

## 额外回归范围

- draw.io Desktop CLI 导出 SVG 后的文字 fallback 检查
- 插件 SVG 清洗结果检查
- PowerPoint 真正插入后的 `pptx` 包内 SVG 内容检查

## 关键观察

- 原始导出 SVG：
  - `OriginalForeignObject=True`
  - `OriginalEmbeddedTextPng=True`
- 清洗后的 SVG：
  - `SanitizedForeignObject=False`
  - `SanitizedEmbeddedTextPng=False`
  - `SanitizedHasTspan=True`
- 写入 PowerPoint 并保存后的 `pptx` 包内 SVG：
  - `StoredSvgContainsForeignObject=False`
  - `StoredSvgContainsEmbeddedTextPng=False`
  - `StoredSvgHasTspan=True`

## 测试结果汇总

- `ReleaseBuild`：`PASS`
  - `MSBuild Rebuild (Release|x64)` 完成，`0 warning / 0 error`
  - `DrawioPpt.Core.dll = 1.0.1.0`
  - `DrawioPpt.PowerPointAddIn.dll = 1.0.1.0`
- `ReleasePackage`：`PASS`
  - 成功生成 `artifacts\releases\v1.0.1\package`
  - 成功生成 `artifacts\releases\v1.0.1\DrawioPpt-v1.0.1.zip`
- `InstallRelease`：`PASS`
  - 成功安装到 `C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1`
  - 注册表 `CodeBase` 指向安装包内 `DrawioPpt.PowerPointAddIn.dll`
- `InstalledUrlSmoke`：`PASS`
  - 安装包内 `url-editor-smoke.ps1 -SkipBuild` 返回 `0`
  - 插件日志记录了 `configure -> init -> load -> save -> export -> DialogResult=OK`
- `SvgVectorTextRegression`：`PASS`
  - 原始 draw.io 导出样本带 `foreignObject + data:image/png`
  - 清洗后与写入 `pptx` 后均不再保留位图文字 fallback，并保留 `tspan`

## 产物与日志

- 发布包清单：`artifacts\releases\v1.0.1\package\PACKAGE.txt`
- 安装验证目录：`C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1`
- URL smoke 插件日志：`C:\Users\nabul\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log`
- 本版本日志目录：`artifacts\logs\v1.0.1\`

## 结论

`v1.0.1` 的核心修复目标已经通过真实 PowerPoint 宿主链路验证：draw.io 导出的文字位图 fallback 会在进入 PowerPoint 前被转换为 SVG 文本，最终保存到 `pptx` 中的仍是矢量文字内容。
