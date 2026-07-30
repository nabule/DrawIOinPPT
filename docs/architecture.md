# 架构设计

[中文](./architecture.md) | [English](./architecture.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 1. 目标

本项目的目标是让 Office 文档中的 Draw.io 图形具备以下能力。当前已覆盖 PowerPoint Desktop，并新增 Word Desktop 宿主：

- PowerPoint 直接插入原始 SVG，保持矢量质量；Word 使用 1240px 等比例 PNG 预览，降低复杂 SVG 的原生布局开销。
- PowerPoint 使用既有编辑入口；Word 必须先选中图片再点击显式命令，才重新进入 Draw.io 编辑。
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

当前共享设置模型包含编辑器模式、桌面版路径、URL 地址、Office 兼容 SVG 标签、选中自动打开、保存后自动刷新、sidecar 保留策略，以及新建或编辑时是否显示图形信息弹窗。其中“选中自动打开”仅由 PowerPoint 使用；Word 为保证拖动热路径零插件处理，不订阅选区变化，也不提供自动打开入口。

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
- Word 显式命令触发的实时图片识别
- Word `InlineShape` / 浮动 `Shape` 对象模型交互
- Word 文档级 `CustomXMLParts` 存储与清理
- 外部编辑器启动与回写编排

当前 Word 宿主复用 PowerPoint 程序集中的公共编辑器、SVG 导出与清理服务，后续可以再把这些公共服务抽成独立 `OfficeShared` 项目。展示策略按宿主区分：PowerPoint 直接插入原始 SVG，按幻灯片边界做 `contain` 等比缩放并居中，不改写 `viewBox`、不向画布补白；Word 从 Draw.io 源生成宽度 1240px 的等比例 PNG，插入为浮动 `Shape`，环绕固定为 `wdWrapFront`。draw.io Desktop 导出失败时，Word 以 1240px 为基准、按源比例生成轻量 PNG 占位预览，极端纵向图仍受 1200px 高度上限保护，不把复杂 SVG 放回 Word；完整图源仍在文档元数据中，用户选中后点击命令仍可编辑。

Word 不订阅 `WindowSelectionChange`，启动时也不读取当前选区、图片属性或 `AlternativeText`。因此选中、拖动、缩放和重新定位完全走 Word 原生路径，插件不在该热路径中执行元数据识别、Ribbon 状态刷新、文档扫描或自动打开。Word 功能区按钮保持可点击；用户先选中图片，再点击“重新编辑”“刷新”“绑定”或“清除绑定”，命令才读取实时选区并校验元数据。双击属于用户显式编辑动作，仍可按需读取当前图片。

### 3.4 安装与注册脚本

职责：

- `scripts\register-office-addins.ps1` 是开发目录和发布包内的统一注册入口，会依次注册 PowerPoint 与 Word 两个当前用户级 COM Add-in。
- `scripts\unregister-office-addins.ps1` 是统一注销入口，会依次清理 PowerPoint 与 Word 注册项。
- `scripts\verify-office-install.ps1` 是安装验收入口，会检查插件文件、当前用户级 COM 注册、`CodeBase` 指向、Office 禁用项，以及可选的 Word/PowerPoint COM 加载和版本自动化回调。
- `scripts\register-addin.ps1`、`scripts\register-word-addin.ps1` 以及对应注销脚本保留为单宿主排障入口。
- 发布包的 `install.cmd` 调用 `scripts\install-release.ps1`，安装脚本复制文件后调用统一注册入口，并立即执行非交互注册验收，避免只安装 PowerPoint 而遗漏 Word。
- 升级安装遵循“先暂存、再清理、后复制”的顺序：`install-release.ps1` 先按 `PACKAGE.txt` 清单把发布包有效载荷复制到 `%TEMP%`，再删除旧安装目录并从暂存目录安装，避免新包位于旧安装目录内时删除安装源，也避免旧残留文件回流。
- Ribbon 的版本标签由 Office 直接调用两个 `Connect` COM 入口类的 `GetVersionSummary`。入口类转发到各自的 `RibbonController` 和共享 `BuildInfo`，确保 Word、PowerPoint 显示的 BuildId 与安装包 `BuildInfo.txt` 一致。安装验收则通过每个宿主的 `COMAddIn.Object` 自动化对象读取同一份宿主版本文本；这与 Ribbon 入口分离，但都最终调用同一个 `AddInHost`。

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

- 文档级 `CustomXMLParts` 是主存储。PowerPoint 和 Word 都在这里保存完整的插件 envelope；envelope 内包含 `formatVersion`、`diagramId`、图形名称、编辑模式、编辑目标、sidecar path、更新时间和压缩后的 `DrawioXml`。
- `DrawioXml` 使用 `gzip + base64` 存储在 envelope 的 `drawioXml` 节点中，减少 Office 文件体积并避免把大段原始 XML 直接塞进图形属性。
- PowerPoint 图形通过 `Shape.Tags["DRAWIO_PPT_ID"]` 和 `Shape.Tags["DRAWIO_PPT_PART_ID"]` 关联到文档级 XML 部件。
- Word 没有等价的 `Shape.Tags`。正常写入时，`InlineShape/Shape.AlternativeText` 只保存不含 `DrawioXml` 的轻量 envelope，字段包括 `formatVersion`、`diagramId`、图形名称、编辑模式、编辑目标、sidecar path 和更新时间；插件据此从 `Document.CustomXMLParts` 读取完整 envelope。
- 如果 Word `CustomXMLParts.Upsert` 失败，插件会把全量 envelope 留在图片 `AlternativeText` 中，避免保存过程中丢失源数据。这个失败回退不改变 PowerPoint 的存储策略。
- 旧版 Word 文档可能只在 `AlternativeText` 中保存全量 envelope。首次读取时，插件会把完整数据迁入 `Document.CustomXMLParts`；写入成功后将图片改写为轻量引用，重复读取不会继续创建新部件。
- SVG 展示文件还会补写 draw.io `content` 元数据，作为第三层恢复信息；它不是主存储，但有助于单个 SVG 被导出或单独排查时恢复源图。
- sidecar `.drawio` 文件现在更接近编辑缓存和人工备份：桌面版 draw.io 需要真实文件路径时使用它，文档内源数据才是跨机器移动时的主依据。

当前没有把 `.drawio` 作为 OLE 对象或 `EmbeddedPackagePart` 附件嵌入 Office 文件。这样做可以减少 Office 安全提示、文件关联依赖和跨宿主差异；代价是插件需要维护“图形引用 -> 文档级 XML 部件”的索引关系。

### 已知边界

- 文档级 `CustomXMLParts` 会跟随整个 `.pptx` / `.docx` 保存和移动，但单独复制一个图形到另一个文档时，Office 不保证对应的文档级 XML 部件也被复制。
- 对 Word 而言，正常图片上的轻量 `AlternativeText` 不含 Draw.io XML；因此只复制单个图片到另一文档后，编辑源可能无法恢复。建议复制整份 `.docx`，或事先保存/保留 sidecar `.drawio` 或 Draw.io XML。只有旧版图片或主存储写入失败的回退图片，`AlternativeText` 才可能仍包含全量数据。
- 如果用户使用文档检查器、企业 DLP、另存旧格式、导出 PDF 或第三方 Office 兼容软件，自定义 XML 部件可能被移除或忽略。
- `AlternativeText` 是可见的辅助说明属性，不适合长期作为大容量主存储。Word 正常路径只把它作为识别引用，写入失败和旧版迁移期间才承担全量回退。
- 清除绑定会移除图形/图片上的插件元数据，并清理无引用的文档级 XML 部件，但不会删除可见图片本身。

## 5. 编辑工作流

### 5.1 新建

1. 点击 Ribbon 的“新建图形”
2. 启动外部 draw.io 编辑器
3. 用户保存 `.drawio` 文件
4. 插件导出 SVG；Word 正常路径再从 Draw.io 源生成 1240px 等比例 PNG 预览
5. PowerPoint 插件把原始 SVG 按 `contain` 等比居中插入当前幻灯片；Word 插件把 PNG 作为 `wdWrapFront` 浮动图片插入当前光标位置
6. 插件写入元数据包和文档级 `CustomXMLParts`
7. 如果 `ShowDiagramInfoDialog` 开启，显示图形名称、Diagram ID、编辑模式和 `.drawio` 工作文件路径等信息

### 5.2 编辑

1. 用户选中一个图形或图片
2. PowerPoint 可按原有选择事件更新状态；Word 选中本身不触发插件处理
3. Word 用户点击“重新编辑”（或双击图片），PowerPoint 用户执行原有编辑入口
4. 插件在显式操作发生后识别对象并读取元数据包
5. 打开本地编辑器或 URL 编辑器
6. 用户保存
7. 插件重新生成展示内容：PowerPoint 仍替换为原始 SVG；Word 重新生成 1240px 等比例 PNG 并替换浮动图片
8. 如果 `ShowDiagramInfoDialog` 开启，进入编辑流程时显示当前图形的信息弹窗

### 5.3 URL 编辑器的 WebView2 用户数据目录

URL 模式在创建 `WebView2` 前显式建立 `CoreWebView2Environment`，并把浏览器用户数据目录固定为：

```text
%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2
```

该目录由当前用户创建，PowerPoint 和 Word 共用。这样 WebView2 不会沿用默认规则，在 `POWERPNT.EXE` 或 `WINWORD.EXE` 可执行文件旁创建 profile；后者通常位于受保护的 Office 安装目录，可能导致 `E_ACCESSDENIED`。目录仅保存 WebView2 的浏览器 profile，不保存 Draw.io 源 XML；图形源数据仍由 Office 文档的 `CustomXMLParts` 与图片元数据管理。

## 6. 更新策略

第一阶段采用“替换展示对象”的方式实现更新，优先保证流程跑通。

保留以下信息：

- 位置
- 尺寸
- 旋转角度
- 图层顺序

PowerPoint 新建时按幻灯片可用边界对原始 SVG 做 `contain` 缩放并居中；替换时保留现有宽度、按 SVG 源比例校正高度，不扩展 SVG `viewBox`，因此不会在图内制造留白。Word 新建、重新编辑和刷新时使用 1240px 等比例 PNG；fallback 占位预览以 1240px 为基准、按源比例生成，极端纵向图受 1200px 高度上限保护。图片始终转换为浮动 `Shape` 并设置 `wdWrapFront`。横向和纵向图均按文档可用边界 `contain`，纵向回归比例为 `0.25`。替换采用“先创建并完整配置新图，再删除原图”的事务式顺序；新图创建或配置失败时清理新图、保留原图。预览临时 XML 删除使用有限重试和延迟，失败目录进入后台延迟清理，提供器启动时还会清理过期目录。

如果后续发现动画、超链接或复杂格式在替换时容易丢失，则进入增强版本：

- 直接替换 PPTX 内部图片部件，减少对象重建

## 7. 风险与应对

### 风险 1：`AlternativeText` 容量与兼容性

应对：

- Word 正常路径只在 `AlternativeText` 中写轻量引用，完整 envelope 由 `Document.CustomXMLParts` 保存。
- `CustomXMLParts` 写入失败时保留全量图片回退，旧版全量元数据在读取时幂等迁移。
- 明确单图跨文档复制的兼容取舍，建议复制整份文档或保留 sidecar/Draw.io XML。

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

