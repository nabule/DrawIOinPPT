# 架构设计

[中文](./architecture.md) | [English](./architecture.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 1. 目标

本项目的目标是让 Office 文档中的 Draw.io 图形具备以下能力。当前已覆盖 PowerPoint Desktop，并新增 Word Desktop 宿主：

- 以 SVG 方式显示，保持矢量质量。
- 在 PowerPoint 中选中图形、在 Word 中选中图片后，可以重新进入 Draw.io 编辑。
- 编辑器可配置为本地桌面程序或 Web URL。
- 原始 Draw.io XML 尽量和 `pptx` / `docx` 文件一起移动。

## 2. 技术路线

本项目选择 **Office 原生 COM Add-in**，当前包含 PowerPoint 和 Word 两个宿主，原因如下：

- 可以可靠监听 Office Desktop 的选择事件。
- 可以直接调用本地进程，适合桌面版 draw.io 联动。
- 可以更自然地处理本地路径、临时文件和后续自动回写。

本项目暂不考虑 WPS，也不把 Office Web Add-in 作为主路线。

## 3. 模块划分

### 3.1 `DrawioPpt.Core`

职责：

- 定义插件设置模型
- 定义 Draw.io 元数据包模型
- 提供元数据序列化和压缩
- 生成 sidecar 文件路径

当前设置模型包含编辑器模式、桌面版路径、URL 地址、Office 兼容 SVG 标签、选中自动打开、保存后自动刷新、sidecar 保留策略，以及新建或编辑时是否显示图形信息弹窗。

### 3.2 `DrawioPpt.PowerPointAddIn`

职责：

- COM Add-in 入口与注册
- Ribbon 命令暴露
- 选中图形监听
- PowerPoint 对象模型交互
- 外部编辑器启动与回写编排

### 3.3 `DrawioPpt.WordAddIn`

职责：

- Word COM Add-in 入口与注册
- Ribbon 命令暴露
- Word 选中图片监听
- Word `InlineShape` / 浮动 `Shape` 对象模型交互
- Word 文档级 `CustomXMLParts` 存储与清理
- 外部编辑器启动与回写编排

当前 Word 宿主复用 PowerPoint 程序集中的公共编辑器和 SVG 服务，后续可以再把这些公共服务抽成独立 `OfficeShared` 项目。

### 3.4 安装与注册脚本

职责：

- `scripts\register-office-addins.ps1` 是开发目录和发布包内的统一注册入口，会依次注册 PowerPoint 与 Word 两个当前用户级 COM Add-in。
- `scripts\unregister-office-addins.ps1` 是统一注销入口，会依次清理 PowerPoint 与 Word 注册项。
- `scripts\register-addin.ps1`、`scripts\register-word-addin.ps1` 以及对应注销脚本保留为单宿主排障入口。
- 发布包的 `install.cmd` 调用 `scripts\install-release.ps1`，安装脚本复制文件后调用统一注册入口，避免只安装 PowerPoint 而遗漏 Word。

## 4. 数据存储策略

### 初始版本策略

为了尽快拿到可用 MVP，初始版本采用两层存储：

PowerPoint：

- 图形标识：`Shape.Tags["DRAWIO_PPT_ID"]`
- 图形元数据包：`Shape.AlternativeText`

Word：

- 图片元数据包：`InlineShape/Shape.AlternativeText`
- 图形名称：`InlineShape/Shape.Title`

`AlternativeText` 中存的是插件自定义 XML 包，内部包含：

- diagram id
- display name
- editor mode
- editor target
- sidecar path
- 更新时间
- 原始 draw.io XML

其中原始 draw.io XML 会先进行 `gzip + base64` 压缩，减少体积。

### 当前嵌入模型

当前 PowerPoint 已增强为 `Presentation.CustomXMLParts + Shape.Tags + Shape.AlternativeText`，Word 已增强为 `Document.CustomXMLParts + InlineShape/Shape.AlternativeText`。

这意味着 Draw.io 源 XML 当前已经嵌入在 `.pptx` / `.docx` 文件内部，而不是只依赖外部 `.drawio` 文件。具体分工如下：

- 文档级 `CustomXMLParts` 是主存储，保存完整的插件 envelope。envelope 内包含 `diagramId`、图形名称、编辑模式、编辑目标、sidecar path、更新时间和压缩后的 `DrawioXml`。
- `DrawioXml` 使用 `gzip + base64` 存储在 envelope 的 `drawioXml` 节点中，减少 Office 文件体积并避免把大段原始 XML 直接塞进图形属性。
- PowerPoint 图形通过 `Shape.Tags["DRAWIO_PPT_ID"]` 和 `Shape.Tags["DRAWIO_PPT_PART_ID"]` 关联到文档级 XML 部件。
- Word 没有等价的 `Shape.Tags`，所以通过图片 `AlternativeText` 中的压缩 envelope 解析 `diagramId`，再从 `Document.CustomXMLParts` 找到最新完整数据。
- `AlternativeText` 仍保留一份兼容回退数据，用于旧文档迁移、复制粘贴后的恢复，以及无法立即读取文档级部件时的降级识别。
- SVG 展示文件还会补写 draw.io `content` 元数据，作为第三层恢复信息；它不是主存储，但有助于单个 SVG 被导出或单独排查时恢复源图。
- sidecar `.drawio` 文件现在更接近编辑缓存和人工备份：桌面版 draw.io 需要真实文件路径时使用它，文档内源数据才是跨机器移动时的主依据。

当前没有把 `.drawio` 作为 OLE 对象或 `EmbeddedPackagePart` 附件嵌入 Office 文件。这样做可以减少 Office 安全提示、文件关联依赖和跨宿主差异；代价是插件需要维护“图形引用 -> 文档级 XML 部件”的索引关系。

### 已知边界

- 文档级 `CustomXMLParts` 会跟随整个 `.pptx` / `.docx` 保存和移动，但单独复制一个图形到另一个文档时，Office 不保证对应的文档级 XML 部件也被复制；因此图形上的 `AlternativeText` 和 SVG `content` 仍然保留恢复价值。
- 如果用户使用文档检查器、企业 DLP、另存旧格式、导出 PDF 或第三方 Office 兼容软件，自定义 XML 部件可能被移除或忽略。
- `AlternativeText` 是可见的辅助说明属性，不适合长期作为大容量主存储；当前只把它作为识别和兼容回退。
- 清除绑定会移除图形/图片上的插件元数据，并清理无引用的文档级 XML 部件，但不会删除可见图片本身。

## 5. 编辑工作流

### 5.1 新建

1. 点击 Ribbon 的“新建图形”
2. 启动外部 draw.io 编辑器
3. 用户保存 `.drawio` 文件
4. 插件导出 SVG
5. PowerPoint 插件将 SVG 插入当前幻灯片；Word 插件将 SVG 插入当前光标位置
6. 插件写入元数据包和文档级 `CustomXMLParts`
7. 如果 `ShowDiagramInfoDialog` 开启，显示图形名称、Diagram ID、编辑模式和 `.drawio` 工作文件路径等信息

### 5.2 编辑

1. 用户选中一个图形或图片
2. 插件监听选择变化
3. 识别是否为插件管理的 Draw.io 图形
4. 读取元数据包
5. 打开本地编辑器或 URL 编辑器
6. 用户保存
7. 插件重新生成 SVG 并替换展示内容
8. 如果 `ShowDiagramInfoDialog` 开启，进入编辑流程时显示当前图形的信息弹窗

## 6. 更新策略

第一阶段采用“替换展示对象”的方式实现更新，优先保证流程跑通。

保留以下信息：

- 位置
- 尺寸
- 旋转角度
- 图层顺序

Word 中 `InlineShape` 是正文流式对象，没有 PowerPoint 那样的固定画布和图层顺序；当前优先保留图片大小、插入位置和浮动图片的环绕/相对位置。

如果后续发现动画、超链接或复杂格式在替换时容易丢失，则进入增强版本：

- 直接替换 PPTX 内部图片部件，减少对象重建

## 7. 风险与应对

### 风险 1：`AlternativeText` 容量与兼容性

应对：

- 先压缩 XML
- 在计划中保留文档级存储升级路线

### 风险 2：外部编辑器保存检测

应对：

- 本地模式优先支持“进程退出后回写”
- 第二阶段增加文件监听和节流

### 风险 3：SVG 替换保真

应对：

- 先保留位置信息与尺寸
- 再针对动画和复杂样式做专项修复

### 风险 4：Word 流式排版与浮动图片锚点

应对：

- MVP 优先保障 `InlineShape` 图片闭环。
- 浮动 `Shape` 替换时保留锚点、大小、环绕方式和相对位置。
- 用真实 Word E2E 验证保存、重开、再编辑链路。

