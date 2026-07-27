# Word 选区事件驱动同步设计

## 目标

移除 Word 插件每 500ms 的选区状态轮询，改为由 Word 的 `WindowSelectionChange` 事件驱动选区状态、功能区刷新和“自动打开”逻辑，避免在图片拖动与缩放期间周期性占用 Word UI 线程。

## 已确认事实

- `SelectionMonitor` 已订阅 `WindowSelectionChange`，可在选区变化时提供 `SelectionContext`。
- `AddInHost` 还启动了 `System.Windows.Forms.Timer`，每 500ms 再读取当前 Word 选区；读取托管图片时会解析图片 `AlternativeText` 中的压缩 Draw.io 元数据。
- 用户已确认 SVG 本身不卡顿，本次不改变 SVG、Draw.io XML、图片插入格式或 Word 浮动图片的环绕策略。

## 方案

1. 删除 `AddInHost` 中 `Timer` 字段、构造、启动、销毁和 Tick 回调。
2. 保留 `SelectionMonitor` 的 `WindowSelectionChange` 订阅；启动时继续执行一次 `ReadLiveSelectionContext()`，确保加载插件时功能区状态正确。
3. 保持现有 `OnSelectionChanged -> ApplySelectionContext(..., true)` 路径，以继续支持功能区刷新和已开启的“自动打开”。
4. “编辑、刷新、绑定、清除绑定”等命令继续在用户点击时读取当前选区，不依赖后台轮询。

## 验收标准

- Word 宿主源码不再创建或启动 `System.Windows.Forms.Timer`。
- `WindowSelectionChange` 仍是唯一后台选区同步入口，且启动时保留一次初始同步。
- 真实 Word URL E2E 仍能完成：新建、选中、编辑、保存、重开和再次回写。
- Debug/Release 构建、发布包、安装验收和 Word COM 加载检查均通过。

## 范围边界

- 不修改 PowerPoint 宿主。
- 不改变 `AlternativeText`/`CustomXMLParts` 的数据兼容策略；其大载荷优化另行处理。
- 不承诺改变 Word 浮动图片“上下型环绕”本身造成的排版开销。
