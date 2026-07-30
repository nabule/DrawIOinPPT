# DrawioPpt

[中文](./README.md) | [English](./README.en.md)

DrawioPpt 是一个面向 Microsoft Office Desktop 的 Draw.io / diagrams.net 原生插件。它让你可以在 PowerPoint 和 Word 里直接插入、识别、重新编辑和刷新 Draw.io 图形，而不是把图形当成一次性截图维护。

当前公开版本：[DrawioPpt v1.0.10](https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.10)

上一公开版本：[DrawioPpt v1.0.9](https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.9)

## 项目价值

DrawioPpt 解决的是 Office 文档里流程图、架构图和说明图的长期维护问题：

- 图形以 SVG 方式插入 Office，显示质量比截图更稳定。
- 选中已有图形后可以重新进入 Draw.io 编辑，保存后回写原图形。
- Draw.io 源 XML 会嵌入 `.pptx` / `.docx` 文件内部，文件移动或发给别人时不必只依赖旁边的 `.drawio` 文件。
- PowerPoint 和 Word 使用同一套编辑器配置，减少重复设置。
- PowerPoint 和 Word 的功能区信息区、共享设置窗口都会显示当前 BuildId，便于确认 Office 实际加载的插件版本。
- 支持桌面 draw.io 和内嵌 URL 编辑器两种模式，适合离线、内网和私有化 diagrams.net 场景。

## 支持的 Office 宿主

### PowerPoint

- 在当前幻灯片插入新的 Draw.io 图形。
- 识别、重新编辑和刷新已绑定图形。
- 支持双击或选中图形后自动打开编辑器。
- 通过 `Presentation.CustomXMLParts + Shape.Tags + AlternativeText` 保存源数据和引用关系。
- 支持 sidecar `.drawio` 工作文件自动重定位。

### Word

- 在当前光标位置插入 Draw.io 图形。
- 选中已绑定图片后重新编辑、刷新、绑定和清除绑定。
- 支持 `InlineShape` 和浮动 `Shape` 图片。
- 通过 `Document.CustomXMLParts + AlternativeText` 保存源数据和轻量引用关系：完整 Draw.io XML 留在文档级主存储，图片属性不再重复承载复杂图形的大 XML。
- Word 与 PowerPoint 共用桌面编辑器、URL 编辑器和基础设置。
- Word 不监听普通选区变化，也不在启动时读取图片元数据；选中、拖动、缩放和重新定位期间不执行插件逻辑。
- 用户先选中图片，再点击“重新编辑、刷新、绑定、清除绑定”；显式命令点击后只读取一次实时选区。

## 核心功能

- `新建`：创建 Draw.io 图形并插入当前 PPT/Word。
- `重新编辑`：打开当前选中的已绑定图形，保存后回写原位置。
- `刷新`：从绑定源数据重新生成 SVG。
- `绑定`：把普通图片纳入插件管理。
- `清除绑定`：移除插件元数据，保留可见图形。
- `自动打开`（仅 PowerPoint）：选中已绑定图形时自动进入编辑流程；Word 为保证拖动热路径零插件开销，不提供此入口。
- `设置`：配置桌面编辑器路径、URL 编辑器地址、sidecar 保存策略和信息弹窗。
- 版本核对：功能区信息区和设置窗口底部显示 `v1.0.10+<git-short-hash>` 形式的 BuildId。

## 源数据如何保存

DrawioPpt 当前采用多层保存策略：

- 主存储：Office 文档级 `CustomXMLParts`，保存压缩后的 Draw.io XML。
- PowerPoint 图形引用：`Shape.Tags` 记录 `diagramId` 和文档级 XML 部件引用。
- Word 图片引用：正常情况下，图片 `AlternativeText` 只保存不含 `DrawioXml` 的轻量 envelope，包含 `formatVersion`、`diagramId`、名称、编辑模式/目标、sidecar 路径和更新时间；完整 envelope 保存于 `Document.CustomXMLParts`。
- Word 失败回退：如果写入 `Document.CustomXMLParts` 失败，图片 `AlternativeText` 会保留全量 envelope，避免源数据丢失；旧版全量图片元数据在首次读取时自动迁入文档级主存储。
- SVG 兜底：生成的 SVG 会补写 draw.io `content` 元数据。
- sidecar `.drawio`：作为桌面编辑缓存、人工备份和自动刷新辅助，不是唯一数据源。

这意味着单独移动 `.pptx` / `.docx` 时，Draw.io 源数据通常会跟随 Office 文件本身一起移动。需要注意的是，文档检查器、企业 DLP、另存旧格式、PDF 导出或第三方 Office 兼容软件可能移除或忽略自定义 XML。

## 推荐 draw.io Desktop

桌面模式推荐使用 [draw.io Desktop v30.0.4 XML 工具修改版](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1)。

推荐资产：

```text
draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip
```

解压后在插件设置里把 `桌面版路径` 指向：

```text
win-unpacked\draw.io.exe
```

这个版本在 draw.io 第二行工具栏新增两个 XML 按钮：

- 从剪贴板粘贴 draw.io XML 源码并显示为图形。
- 将当前画布复制为 draw.io XML 源码。

这对排查 Office 文件内嵌 XML、手工恢复图形和对比插件回写内容很有用。

## 快速安装

1. 从 [v1.0.10 Release](https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.10) 下载 `DrawioPpt-v1.0.10.zip`。
2. 解压到本地目录。
3. 关闭 PowerPoint 和 Word。
4. 双击 `install.cmd`。
5. 打开 PowerPoint 或 Word，确认出现 `Draw.io` 功能区。
6. 在功能区信息区或设置窗口底部核对 BuildId；发布包安装脚本也会输出同一个 BuildId。

已安装过旧版本时，重复运行新发布包里的 `install.cmd` 就是升级；不要运行旧安装目录里遗留的旧 `install.cmd`。若安装失败，命令窗口会保留 PowerShell 原始错误并打印退出码，按提示关闭 Word、PowerPoint 后再运行新包。功能区的版本文本应与安装程序输出的 BuildId 一致。

卸载时关闭 PowerPoint 和 Word，然后运行 `uninstall.cmd`。

发布包安装脚本会同时安装 PowerPoint 和 Word 插件。仓库开发调试时，先构建再运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-office-addins.ps1
```

如果只需要排查单个宿主，可以分别运行 `scripts\register-addin.ps1` 或 `scripts\register-word-addin.ps1`。

安装后可以运行下面的验收脚本，确认 Word 和 PowerPoint 都已注册并能被 Office COM 加载：

```powershell
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

发布包根目录和 `bin` 目录会包含 `BuildInfo.txt`，`PACKAGE.txt` 也会记录 `BuildId` 和 `GitShortHash`，用于排查“Office 加载的 DLL 是否来自当前发布包”。

## 文档入口

- [用户手册](./docs/user-guide.md)
- [安装与联调说明](./docs/installation.md)
- [Word 插件设计与使用说明](./docs/word-addin.md)
- [架构设计](./docs/architecture.md)
- [回归清单](./docs/regression-checklist.md)
- [v1.0.10 发布说明](./docs/release-notes-v1.0.10.md)
- [v1.0.10 已验证 E2E 测试报告](./docs/e2e-test-report-v1.0.10.md)
- [v1.0.10 发布证据与日志索引](./docs/release-evidence-v1.0.10.md)

开发计划、历史发布记录和优化 backlog 也放在 [docs](./docs/) 目录中，README 只保留项目定位和上手入口。

## 技术边界

- 目标环境是 Windows + Microsoft PowerPoint/Word Desktop。
- 当前不支持 Office Web、WPS 或移动端 Office。
- 编辑能力依赖本插件；没有安装插件的一方仍能看到图形，但不能直接进入 Draw.io 回写流程。
- 发布包使用当前用户级 COM 注册，不需要管理员权限。
