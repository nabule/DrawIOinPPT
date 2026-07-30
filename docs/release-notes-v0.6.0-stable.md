# v0.6.0-stable

[中文首页](../README.md) | [English Overview](../README.en.md)

这是当前仓库的首个稳定收口版本，目标是把 Draw.io 与 PowerPoint 的核心闭环做完整，并补齐发布前需要的稳定化与文档交付物。

## 这版先解决什么问题

- PowerPoint 里的 Draw.io 图以前更像一次性图片，后续修改不方便、容易重新截图替换。
- 图形源文件和 PPT 容易分离，换目录、发给别人或长期维护时难以找回可编辑源。
- 桌面 draw.io 和 URL 编辑器两条链路需要统一到一个可安装、可测试的稳定插件流程里。
- 旧版图形数据需要兼容和清理能力，避免升级后已有 PPT 里的图形无法继续维护。
- 安装、异常诊断和回归验证资料需要更完整，降低交付和排障成本。

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
