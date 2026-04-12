# v0.6.0-stable

[中文首页](../README.md) | [English Overview](../README.en.md)

这是当前仓库的首个稳定收口版本，目标是把 Draw.io 与 PowerPoint 的核心闭环做完整，并补齐发布前需要的稳定化与文档交付物。

## 本版本包含

- PowerPoint Desktop 原生 COM Add-in 宿主
- Ribbon 菜单与图形选择识别
- SVG 插入、替换与基础属性保真
- 本地 `draw.io` / `diagrams.net Desktop` 编辑链路
- URL 模式编辑链路：`WebView2 + diagrams.net embed`
- `Presentation.CustomXMLParts` 文档级源数据存储
- `Shape.Tags` 与 `AlternativeText` 回退兼容
- 旧版 `AlternativeText` 数据自动迁移
- 孤儿 `CustomXMLPart` 自动清理
- SVG 文件内嵌 draw.io `content` 元数据
- URL 模式日志、超时诊断与 `Test URL`
- 统一异常收口、安装说明和回归清单

## 已完成的里程碑

- `M0` 初始骨架
- `M1` 元数据 MVP
- `M2` SVG 插入与替换
- `M3` 本地 draw.io 模式
- `M4` URL 模式
- `M5` 自动更新与稳定化

## 推荐验证项

1. 运行 `scripts/dev-check.ps1`
2. 运行 `scripts/build.ps1`
3. 运行 `scripts/url-editor-smoke.ps1`
4. 使用 `scripts/register-addin.ps1` 注册插件
5. 在 PowerPoint 中分别验证桌面模式和 URL 模式

## 相关文档

- [安装与联调说明](./installation.md)
- [回归清单](./regression-checklist.md)
- [详细开发计划](./development-plan.md)
