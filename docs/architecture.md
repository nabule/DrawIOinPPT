# 架构设计

[中文](./architecture.md) | [English](./architecture.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 1. 目标

本项目的目标是让 PowerPoint 中的 Draw.io 图形具备以下能力：

- 以 SVG 方式显示，保持矢量质量。
- 选中后可以重新进入 Draw.io 编辑。
- 编辑器可配置为本地桌面程序或 Web URL。
- 原始 Draw.io XML 尽量和 PPT 文件一起移动。

## 2. 技术路线

本项目选择 **PowerPoint 原生 COM Add-in**，原因如下：

- 可以可靠监听 PowerPoint 的图形选择事件。
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

## 4. 数据存储策略

### 初始版本策略

为了尽快拿到可用 MVP，初始版本采用两层存储：

- 图形标识：`Shape.Tags["DRAWIO_PPT_ID"]`
- 图形元数据包：`Shape.AlternativeText`

`AlternativeText` 中存的是插件自定义 XML 包，内部包含：

- diagram id
- display name
- editor mode
- editor target
- sidecar path
- 更新时间
- 原始 draw.io XML

其中原始 draw.io XML 会先进行 `gzip + base64` 压缩，减少体积。

### 后续增强策略

如果后续验证发现 `AlternativeText` 容量或兼容性不足，则进入二阶段增强：

- 将完整 XML 迁移到文档级自定义部件或 OOXML 附加部件
- 图形上仅保留 diagram id 和必要摘要信息

## 5. 编辑工作流

### 5.1 新建

1. 点击 Ribbon 的“新建图形”
2. 启动外部 draw.io 编辑器
3. 用户保存 `.drawio` 文件
4. 插件导出 SVG
5. 插件将 SVG 插入当前幻灯片
6. 插件写入 `DRAWIO_PPT_ID` 和元数据包
7. 如果 `ShowDiagramInfoDialog` 开启，显示图形名称、Diagram ID、编辑模式和 `.drawio` 工作文件路径等信息

### 5.2 编辑

1. 用户选中一个图形
2. 插件监听选择变化
3. 识别是否为插件管理的 Draw.io 图形
4. 读取元数据包
5. 打开本地编辑器或 URL 编辑器
6. 用户保存
7. 插件重新生成 SVG 并替换展示内容
8. 如果 `ShowDiagramInfoDialog` 开启，进入编辑流程时显示当前图形的信息弹窗

## 6. 更新策略

第一阶段采用“替换 shape 内容”的方式实现更新，优先保证流程跑通。

保留以下信息：

- 位置
- 尺寸
- 旋转角度
- 图层顺序

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

