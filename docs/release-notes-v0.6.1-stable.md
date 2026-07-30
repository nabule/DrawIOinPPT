# v0.6.1-stable

[中文首页](../README.md) | [English Overview](../README.en.md)

这是 `v0.6.0-stable` 之后的一个体验优化版本，重点收口 PowerPoint 中 SVG 标签缩放清晰度的问题，并把对应能力接入到插件的 URL 模式工作流。

## 这版先解决什么问题

- URL 模式导出的图形放大后，文字边缘可能发虚，影响演示和正式材料观感。
- URL 编辑器的加载、保存、导出握手不够明确，排查失败时不容易判断卡在哪一步。
- 用户需要知道桌面模式和 URL 模式里如何保持文字更清晰，而不是只看到内部配置项。

## 本版本新增

- URL 模式自动启用 draw.io `simpleLabels`
- URL embed 自动追加 `configure=1` 并处理 `configure -> init -> load -> save -> export` 握手
- 设置页新增 `Use MS Office-compatible SVG labels`
- 默认新建图形继续使用标准空白 Draw.io XML，兼顾桌面导出稳定性
- 文档补充了桌面模式下手工开启 `simpleLabels` 的说明

## 重点改进

- PowerPoint 中缩放 URL 模式导出的 SVG 时，标签边缘通常会比之前更干净
- URL 模式的本地烟雾测试和完整 E2E 测试都已覆盖 `configure` 配置握手
- 发布包文档继续随版本更新，保持安装说明、用户手册和测试报告一致

## 兼容性说明

- `simpleLabels` 目前由插件自动应用在 URL 模式
- Desktop 模式调用外部 draw.io Desktop，插件不会直接改写外部编辑器配置
- 如果需要桌面模式也获得同样效果，建议在 draw.io Desktop 中手动设置：
  - `Extras > Configuration`
  - `{ "simpleLabels": true }`
  - `Apply`

## 相关文档

- [安装与联调说明](./installation.md)
- [用户手册](./user-guide.md)
- [待优化功能点](./optimization-backlog.md)
