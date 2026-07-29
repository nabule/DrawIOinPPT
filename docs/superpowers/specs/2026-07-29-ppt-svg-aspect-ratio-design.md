# Word 与 PPT SVG 原始比例与无留白设计

## 目标

Word 和 PowerPoint 中新建、重新编辑和刷新 Draw.io SVG 时，禁止为了填满固定图片框而扩展 SVG `viewBox`。图片必须按源 SVG 的宽高比显示，不拉伸、不在 SVG 画布内补白。

## 范围与边界

- 同时修改 Word 的 `WordSvgPictureService` 和 PowerPoint 的 `PowerPointSvgShapeService`。
- 保留图片名称、标题、可见性、旋转、层级、动画和点击/悬停动作等既有元数据恢复逻辑。
- 不采用裁剪填充：流程图和架构图不得因消除留白而丢失边缘内容。

## 设计

### Word 新建与重新编辑

1. 新建时保持当前默认宽度，用源 SVG 宽高比计算高度；直接插入原始 SVG。
2. 重新编辑或刷新时保留旧图片宽度，按源比例重算高度；浮动图片保持垂直中心，行内图片仍由 Word 文本流定位。
3. 新图片锁定比例；不调用 `PrepareForPresentation`，因此不会改写源 SVG 的 `viewBox` 或制造内部留白。

### PowerPoint 新建

1. 从源 SVG 读取宽高比。
2. 保持当前默认宽度（幻灯片宽度的 50%）。
3. 用 `height = width / sourceAspectRatio` 计算高度，并在幻灯片中央放置图片。
4. 直接插入原始 SVG，不调用 `PrepareForPresentation`，因此不会改写源 SVG 的 `viewBox` 或制造内部留白。

### PowerPoint 重新编辑或刷新

1. 捕获旧图片的位置、宽度和元数据。
2. 使用源 SVG 宽高比，保留旧图片宽度，重新计算高度。
3. 按旧图片的垂直中心调整 `Top`，避免高度校正后图片整体跳动。
4. 直接插入原始 SVG，恢复既有元数据、动画和层级；新图片锁定比例。

### 异常回退

若 SVG 缺少可解析尺寸，则维持原有几何尺寸并直接插入原始 SVG。不会引入白边补偿，也不会裁剪图形。

## 验收

真实 Word 与 PowerPoint COM E2E 使用 1200×600（2:1）SVG：

- 两端新建图片比例均为 2:1；
- 两端将旧图模拟为 400×400 后重新编辑，结果仍为 2:1；
- 两端重编辑后宽度仍为 400、高度为 200；浮动图片的垂直中心不变；
- SVG 源文件的 `viewBox` 未被适配逻辑改写；
- 现有完整 Office E2E 与安装态回归均通过。
