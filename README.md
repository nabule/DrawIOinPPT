# DrawioPpt

PowerPoint Desktop 原生插件项目，用于在 PPT 中插入、识别、编辑和更新 Draw.io 图形。

当前仓库处于 `v0.1.0-initial` 阶段，重点完成了以下内容：

- 明确技术路线：PowerPoint 原生 COM Add-in，先不考虑 WPS。
- 固化模块边界：宿主层、元数据层、外部编辑器层、文件与路径层。
- 建立可继续开发的代码骨架：解决方案、项目、Ribbon、选中图形监听、设置与元数据模型。
- 建立后续迭代计划：从元数据 MVP 到外部编辑器联动，再到 SVG 插入与自动更新。

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
  - 配置化 URL 模式

## 构建前提

本仓库当前按 .NET Framework 经典项目组织，优先保证 PowerPoint 宿主兼容性。

建议开发环境：

- Windows
- Microsoft PowerPoint Desktop
- Visual Studio 2022 或等效 MSBuild 环境

当前机器可用的本地工具路径：

- `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe`
- `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe`

## 本地验证

先检查环境：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
```

构建：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

注册插件：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-addin.ps1
```

卸载插件：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister-addin.ps1
```

## 下一步

按照 [docs/development-plan.md](docs/development-plan.md) 推进，优先完成：

1. 元数据写入与读取 MVP
2. SVG 插入与更新
3. 本地 draw.io 进程联动
4. 自动保存监听与回写

