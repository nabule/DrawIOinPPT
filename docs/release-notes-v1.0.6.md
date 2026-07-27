# v1.0.6

[中文](./release-notes-v1.0.6.md) | [English](./release-notes-v1.0.6.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.6` 优化 Word 中 draw.io 图片移动、调整大小和位置时的响应。SVG、Draw.io XML、图片格式、`AlternativeText`/`CustomXMLParts` 存储策略及浮动图片布局策略均未改变。

## 主要变化

- 删除 Word 宿主每 500ms 读取选区的 `System.Windows.Forms.Timer`，不再周期性解析图片元数据。
- 保留 `WindowSelectionChange` 和插件启动时的一次选区读取，用于功能区状态与已开启的自动打开逻辑。
- 编辑、刷新、绑定、清除绑定在用户点击时始终读取当前 Word 选区；当 Word 未立即更新选区状态显示时，手动命令仍可正确作用于已选图片。
- 新增无轮询源码约束、真实 Word COM 浮动图片选择测试，并将无轮询约束接入 GitHub Actions。
- 修复 Word URL E2E：mock 服务改为测试进程内自托管并做 HTTP 可达性检查；测试期间隔离已注册插件并恢复原加载设置，避免双 Host 干扰。
- 发布包包含 `word-selection-event-e2e.ps1`，可对发布 DLL 运行真实 Word COM 验证。

## 使用提示

Word 的浮动图片布局本身仍由 Word 处理；本次只移除了插件每半秒一次的后台选区访问。若功能区状态没有立即刷新，选中图片后直接点击“重新编辑”“刷新”“绑定”或“清除绑定”即可，命令会现场读取选区。

执行 `word-url-e2e.ps1` 或 `word-selection-event-e2e.ps1` 前请关闭 Word，且不要并行运行其他 Word 自动化。详细验证见 [v1.0.6 E2E 测试报告](./e2e-test-report-v1.0.6.md)。
