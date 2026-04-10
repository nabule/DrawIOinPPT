# DrawioPpt

PowerPoint Desktop 原生插件项目，用于在 PPT 中插入、识别、编辑和更新 Draw.io 图形。

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

## 目录结构

- `docs/`：架构、计划和初始版本说明。
- `scripts/`：环境检查、构建、注册与卸载脚本。
- `src/DrawioPpt.Core/`：与宿主无关的核心模型、序列化和设置存储。
- `src/DrawioPpt.PowerPointAddIn/`：PowerPoint COM Add-in 宿主骨架。

## 当前技术选择

- 宿主：C# + PowerPoint COM Add-in
- Office 接口：`Microsoft.Office.Interop.PowerPoint`
- Ribbon：`Microsoft.Office.Core.IRibbonExtensibility`
- 初始元数据存储策略：`Shape.Tags + Shape.AlternativeText`
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

注册插件：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-addin.ps1
```

卸载插件：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister-addin.ps1
```

## 下一步

按照 [docs/development-plan.md](docs/development-plan.md) 继续推进，优先完成：

1. 实机验证 URL 模式的 save/export 消息链路
2. 提升替换图形时对动画、超链接和层级的保真
3. 增加保存日志与错误追踪
4. 评估从 `AlternativeText` 迁移到更稳的文档级存储方案
