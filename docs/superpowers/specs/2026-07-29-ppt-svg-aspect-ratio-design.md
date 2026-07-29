# Word PNG 预览与 PPT SVG 原始比例、无留白设计

## 目标

Word 和 PowerPoint 中新建、重新编辑和刷新 Draw.io 图形时，禁止为了填满固定图片框而扩展 SVG `viewBox` 或强制 1:1。PowerPoint 保留原始 SVG；Word 为降低复杂 SVG 的交互重绘成本，从文档中的 Draw.io XML 导出 1240 像素宽、零边框的等比例 PNG 预览。两端都不得拉伸、裁剪或制造适配留白。

## 范围与边界

- 同时修改 Word 的预览导出、图片几何和 PowerPoint 的 `PowerPointSvgShapeService`。
- 保留图片名称、标题、可见性、旋转、层级、动画和点击/悬停动作等既有元数据恢复逻辑。
- 不采用裁剪填充：流程图和架构图不得因消除留白而丢失边缘内容。

## 设计

### Word 新建与重新编辑

1. `WordPreviewImageProvider` 把文档主存储中的 Draw.io XML 写入隔离临时文件，并调用 draw.io Desktop 以 `1240px` 宽、`border=0` 导出 PNG；失败时使用按源比例生成的轻量 PNG 占位预览，不回退复杂 SVG。
2. 新建时保持当前默认展示宽度，用 PNG/SVG 源宽高比计算高度；转为浮动 `Shape`，锁定比例并使用 `wdWrapFront`，避免正文重排热路径。
3. 重新编辑或刷新时保留旧图片宽度，按新预览源比例重算高度；浮动图片保持垂直中心并统一为 `wdWrapFront`。
4. 图片嵌入 Word 后立即删除本次临时 PNG 和临时 `.drawio`；可编辑 XML 继续存于 `Document.CustomXMLParts`，选中、拖动和缩放期间不读取元数据。

### PowerPoint 新建

1. 从源 SVG 读取宽高比。
2. 使用幻灯片宽度 50%、高度 38% 组成默认边界框。
3. 在该边界框内按源比例 contain 缩放并居中，竖图和横图均不得越界。
4. 直接插入原始 SVG，不调用 `PrepareForPresentation`，因此不会改写源 SVG 的 `viewBox` 或制造内部留白。

### PowerPoint 重新编辑或刷新

1. 捕获旧图片的位置、宽度和元数据。
2. 使用源 SVG 宽高比，保留旧图片宽度，重新计算高度。
3. 按旧图片的垂直中心调整 `Top`，避免高度校正后图片整体跳动。
4. 直接插入原始 SVG，恢复既有元数据、动画和层级；新图片锁定比例。

### 异常回退

PowerPoint 若 SVG 缺少可解析尺寸，则维持原有几何尺寸并直接插入原始 SVG。Word 若 PNG 导出不可用，则按 1240px 基准生成保持源比例的轻量 PNG 占位预览；极端纵向图高度限制为 1200px。两种回退都不引入白边补偿或裁剪。

## 验收

真实 Word 与 PowerPoint COM E2E 使用 1200×600（2:1）输入：

- Word 的 2:1 PNG 预览和 PowerPoint 的 2:1 SVG 新建图片比例均为 2:1；
- 两端将旧图模拟为 400×400 后重新编辑，结果仍为 2:1；
- 两端重编辑后宽度仍为 400、高度为 200；浮动图片的垂直中心不变；
- Word 图片为浮动 `Shape`、`wdWrapFront`，PowerPoint 新建图在默认宽高边界内 contain 且居中；
- Word 预览提供器真实导出宽度 1240px 的等比例 PNG，临时源和 PNG 清理完成；
- 完整 Office E2E 12/12 与安装态回归均通过；
- 用户附件富文档压力测试据实记录 `PointerInputCovered=false`：Windows 桌面接口返回 `GetCursorPos` 访问拒绝和 `SetIsBorderRequired` 不支持，未用 COM 指标冒充真实指针验收。
