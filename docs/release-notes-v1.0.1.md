# v1.0.1

[中文](./release-notes-v1.0.1.md) | [English](./release-notes-v1.0.1.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.1` 是一个面向发布链路和 SVG 显示保真的维护版本，重点收口“draw.io 编辑后回写到 PPT 中的文字应保持矢量清晰”这一问题，并修复本地构建/安装脚本在当前机器上的可执行性。

## 这版先解决什么问题

- draw.io 编辑回写后，PPT 中的文字有时像图片一样发糊，放大后不适合正式演示。
- 本地构建、打包、安装和 E2E 脚本在部分 Windows 环境会因为命令入口差异中断。
- URL 模式测试依赖外部环境时，失败原因不够集中，交付前复测不稳定。

## 本版本重点

### 1. 修复 draw.io 文字在 PPT 中发糊的问题

- draw.io Desktop 导出的部分 SVG 文字标签会带 `foreignObject + data:image/png` 位图 fallback
- 插件现在会在 SVG 兼容清洗阶段把这类文字 fallback 尽量转换为真正的 SVG `<text>/<tspan>`
- 回写到 PowerPoint 之后，PPT 包内保存的 SVG 不再保留这类文字位图 fallback，放大缩小时的清晰度更稳定

### 2. 继续保持 PowerPoint 插入链路为 SVG

- 插件仍然通过 SVG 插入/替换 PowerPoint 图形
- 新增的兼容处理位于 SVG 清洗层，不改变现有 metadata、sidecar、URL 模式和桌面模式的主工作流
- 对既有图形的几何位置、尺寸、旋转和基础动作迁移逻辑没有引入新的发布层改动

### 3. 修复本地发布脚本在当前机器上的执行问题

- `build.ps1`、`package-release.ps1`、`install-release.ps1`、`uninstall-release.ps1`
- `url-editor-smoke.ps1`、`full-e2e-test.ps1`
- 这些脚本内部的二次调用已统一改为 `powershell.exe`，避免因为 `powershell` 别名缺失导致本地构建或安装流程中断
- URL smoke 和完整 E2E 使用的本地 mock 编辑器服务改为脚本内置 `HttpListener`，不再依赖外部 Python

## 升级建议

- 如果你当前机器已经注册过旧版插件，重新运行一次注册或安装流程即可切到 `v1.0.1`
- 如果你之前遇到“draw.io 编辑后在 PPT 中放大文字变糊”，建议优先用 `v1.0.1` 重新编辑并回写一次同一图形
- 如果你通过脚本做本地打包或完整 E2E，建议同步使用当前仓库中的最新脚本

## 配套文档

- [v1.0.1 E2E 测试报告](./e2e-test-report-v1.0.1.md)
- [v1.0.1 发布证据与日志索引](./release-evidence-v1.0.1.md)
- [安装与联调说明](./installation.md)
- [用户手册](./user-guide.md)
