# DrawioPpt

[中文](./README.md) | [English](./README.en.md)

PowerPoint Desktop 原生插件项目，用于在 PPT 中插入、识别、编辑和更新 Draw.io 图形。

当前发布基线：`v1.0.0`

当前仓库已经完成初始骨架，并推进到了可实机联调的阶段，重点完成了以下内容：

- 明确技术路线：PowerPoint 原生 COM Add-in，先不考虑 WPS。
- 固化模块边界：宿主层、元数据层、外部编辑器层、文件与路径层。
- 建立可继续开发的代码骨架：解决方案、项目、Ribbon、选中图形监听、设置与元数据模型。
- 建立后续迭代计划：从元数据 MVP 到外部编辑器联动，再到 SVG 插入与自动更新。
- 已支持给选中的现有图形写入和清除 Draw.io 元数据。
- 已支持识别带元数据的图形，并在 Ribbon 上反馈状态。
- 已支持新建 SVG 预览图形，并将其插入当前幻灯片。
- 已支持刷新已绑定图形，并在桌面模式下监控 `.drawio` 文件变化后自动更新。
- 已支持双击已绑定图形直接打开外部编辑器。
- 已支持 `AutoOpenOnSelection` 设置，选中已绑定图形时可自动进入编辑。
- 已支持自动探测本机已安装的 `draw.io` / `diagrams.net` 桌面版路径。
- 已验证 `draw.io Desktop` CLI 导出 SVG，并收紧了导出参数。
- 已支持当前用户级插件注册，无需管理员权限。
- 已接入 URL 模式编辑器首版：`WebView2 + diagrams.net embed`，保存后会把 SVG 和 XML 回写到 PPT 图形。
- 已在 URL 模式自动启用 `simpleLabels` 配置，提升 PowerPoint 中缩放 SVG 标签时的边缘清晰度。
- 已把 Draw.io 源数据同步写入 `Presentation.CustomXMLParts`，图形保留 `diagramId/customXmlPartId` 引用，并继续兼容旧的 `AlternativeText` 回退。
- 已支持旧版 `AlternativeText` 元数据在读取时自动补写到 `CustomXMLParts`。
- 已支持低频自动清理没有任何图形引用的孤儿 `CustomXMLPart`，减少文档膨胀。
- 已为 URL 模式补充本地日志、初始化超时和导出超时诊断。
- 已统一为桌面导出、URL 导出和预览回退生成的 SVG 补写 draw.io `content` 元数据。
- 设置窗口已支持 `Test URL`，可在保存配置前先验证当前编辑器地址。
- 已重构 PowerPoint Ribbon 分组、状态展示和按钮图标，支持更清晰的模式/对象状态反馈。
- 已支持 sidecar `.drawio` 在 `PPT` 另存为、路径变化和跨机器迁移后按当前文档路径自动重定位。

## 目录结构

- `docs/`：架构、计划和初始版本说明。
- `scripts/`：环境检查、构建、注册与卸载脚本。
- `src/DrawioPpt.Core/`：与宿主无关的核心模型、序列化和设置存储。
- `src/DrawioPpt.PowerPointAddIn/`：PowerPoint COM Add-in 宿主骨架。

关键文档：

- [安装与联调说明](./docs/installation.md)
- [用户手册](./docs/user-guide.md)
- [回归清单](./docs/regression-checklist.md)
- [详细开发计划](./docs/development-plan.md)
- [完整 E2E 测试报告](./docs/e2e-test-report-v1.0.0.md)
- [v1.0.0 发布说明](./docs/release-notes-v1.0.0.md)
- [v1.0.0 发布证据与日志索引](./docs/release-evidence-v1.0.0.md)
- [待优化功能点](./docs/optimization-backlog.md)

## 当前技术选择

- 宿主：C# + PowerPoint COM Add-in
- Office 接口：`Microsoft.Office.Interop.PowerPoint`
- Ribbon：`Microsoft.Office.Core.IRibbonExtensibility`
- 当前元数据存储策略：`Presentation.CustomXMLParts + Shape.Tags + Shape.AlternativeText 回退兼容`
- 外部编辑器模式：
  - 本地桌面版 draw.io/diagrams.net
  - 配置化 URL 模式（`WebView2`）

## 构建前提

本仓库当前按 .NET Framework 经典项目组织，优先保证 PowerPoint 宿主兼容性。

建议开发环境：

- Windows
- Microsoft PowerPoint Desktop
- Visual Studio 2022 或等效 MSBuild 环境

当前机器可用的本地工具路径：

- `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe`
- `C:\Program Files\draw.io\draw.io.exe`
- WebView2 Runtime：已检测

## 本地验证

先检查环境：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
```

构建：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

构建脚本会自动恢复 `Microsoft.Web.WebView2` 包，并默认按 `x64` 构建，与本机 PowerPoint x64 对齐。

URL 模式烟雾测试：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1
```

发布打包：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.0
```

完整真实全流程测试：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.0
```

日志文件默认写入：

```text
C:\Users\<你的用户名>\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

发布测试日志与打包记录会额外保存在：

```text
artifacts\logs\v1.0.0\
```

注册插件：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-addin.ps1
```

卸载插件：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister-addin.ps1
```

## 当前里程碑

- `M0` 已完成
- `M1` 已完成
- `M2` 已完成
- `M3` 已完成
- `M4` 已完成
- `M5` 已完成
