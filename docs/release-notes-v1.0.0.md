# v1.0.0

[中文](./release-notes-v1.0.0.md) | [English](./release-notes-v1.0.0.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

这是 DrawioPpt 的首个 `1.0` 稳定发布版本，目标是把“可编辑 Draw.io 图形跟随 PowerPoint 一起长期维护”的主流程收口到可交付状态，并把安装、测试、日志和发布材料同步完善。

## 这版先解决什么问题

- 用户需要一个真正可交付的 PowerPoint 插件版本，而不是只适合开发调试的半成品。
- PPT 保存、另存为、移动目录或换机器后，旁边的 `.drawio` 工作文件路径容易失效。
- 初次使用时不知道当前选中的图能不能编辑、当前是什么模式、下一步该点哪个按钮。
- 发布包、日志和测试证据需要放在一起，方便安装、复测和问题追溯。

## 版本定位

- 面向 PowerPoint Desktop 的稳定发布版本
- 继续保持 `C# + COM Add-in + WebView2 + draw.io Desktop/URL` 双模式架构
- 以“源码随 PPT 保存、桌面模式可长期维护、URL 模式可实机回写”为核心交付目标

## 本版本重点

### 1. 功能区体验升级

- Ribbon 重组为 `创建 / 当前图形 / 工作流 / 信息`
- 主操作按钮会根据选区状态在 `编辑 / 重新编辑` 之间切换
- 新增更明确的选区状态、模式状态和工作流提示
- 为主要按钮补充自绘图标，降低初次上手成本

### 2. sidecar 路径自动重定位

- 当 PPT 尚未保存时，sidecar 仍会暂存到系统临时目录
- 当 PPT 之后保存、另存为或移动目录后，插件会在下一次编辑时按当前 PPT 路径重建 sidecar
- 将 `pptx` 拷贝到另一台机器后，只要插件和编辑器环境准备好，桌面模式会优先使用当前机器上的新路径

### 3. 数据持久化能力继续收口

- Draw.io 源 XML 继续写入 `Presentation.CustomXMLParts`
- 图形保持 `diagramId / customXmlPartId` 引用
- 兼容旧版 `AlternativeText` 回退数据，并在读取时自动迁移
- SVG 继续补写 draw.io `content` 元数据，便于恢复与排查

### 4. 发布材料更完整

- 发布包继续包含安装脚本、卸载脚本和用户文档
- 新增发布证据文档，用于汇总构建、打包、E2E 和日志产物
- 发布包支持带出当前版本的测试日志，便于交付和追溯

## 兼容性说明

- 支持 `桌面版` 和 `网页地址` 两种编辑模式
- 继续要求 Windows + PowerPoint Desktop x64 + .NET Framework 4.8 + WebView2 Runtime
- URL 模式继续依赖目标地址支持 draw.io embed 协议
- Desktop 模式仍建议本机安装 `draw.io Desktop` 或 `diagrams.net Desktop`

## 升级建议

- 如果你是从 `v0.6.1-stable` 升级，建议重新安装 `v1.0.0` 发布包
- 若旧文档中的 sidecar 曾指向临时目录或旧机器目录，`v1.0.0` 会在下一次编辑时优先切回当前 PPT 路径
- 升级后建议至少执行一次：
  - 新建图形
  - 保存 PPT
  - 关闭并重新打开 PPT
  - 再次编辑同一图形

## 配套文档

- [安装与联调说明](./installation.md)
- [用户手册](./user-guide.md)
- [v1.0.0 E2E 测试报告](./e2e-test-report-v1.0.0.md)
- [v1.0.0 发布证据与日志索引](./release-evidence-v1.0.0.md)
