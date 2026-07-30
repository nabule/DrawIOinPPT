# v1.0.3

[中文](./release-notes-v1.0.3.md) | [English](./release-notes-v1.0.3.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.3` 是面向 Word 插件交付和桌面编辑器配置说明的发布版本。它收口了 Word COM Add-in、Office 文件内嵌 Draw.io XML 的说明，并推荐配合使用支持 XML 剪贴板工具的 draw.io Desktop 修改版。

## 这版先解决什么问题

- 用户希望在 Word 里也能维护 Draw.io 图形，而不是只能在 PowerPoint 中使用插件。
- 交付 `.pptx` / `.docx` 时，用户需要知道源 XML 是否真的跟着 Office 文件走。
- 排查或恢复图形时，需要一个更方便的方式把 Office 内嵌 XML 放回 draw.io Desktop 查看。
- 发布包不应夹带调试符号、历史日志等容易暴露本机路径的信息。

## 主要变化

- 新增 Word Desktop COM Add-in 交付说明，覆盖新建、重新编辑、刷新、绑定和清除绑定 Draw.io 图片。
- 明确 Draw.io 源 XML 已嵌入 `.pptx` / `.docx` 的文档级 `CustomXMLParts`，sidecar `.drawio` 只是桌面编辑缓存、人工备份和自动刷新辅助。
- README、用户手册、安装说明和 Word 专项说明均加入 [draw.io Desktop v30.0.4 XML 工具修改版](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1) 配置建议。
- 发布包默认不夹带 PDB 和历史日志，避免本机路径或调试信息进入分发包。

## 推荐 draw.io Desktop 版本

推荐下载：

```text
draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip
```

下载地址：

```text
https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1
```

SHA256：

```text
B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D
```

这是未签名的 Windows x64 免安装目录版。解压后在插件设置中把 `Desktop Path` / `桌面版路径` 指向：

```text
win-unpacked\draw.io.exe
```

不要只拷贝单独的 exe。

## XML 工具按钮

该 draw.io Desktop 修改版基于官方 `v30.0.4`，第二行工具栏新增两个彩色 XML 图标按钮：

- 从剪贴板粘贴 draw.io XML 源码并显示为图形。
- 将当前画布复制为 draw.io XML 源码。

这对 DrawioPpt 的桌面模式和排障很有帮助：可以直接把 Office 文档内嵌的 Draw.io XML 放回 draw.io Desktop 检查，也可以把当前画布复制成 XML 用于比对、恢复或测试。

## 验证

- Release 构建通过。
- 发布包生成通过。
- 发布包临时安装与卸载通过。
- 包内容检查确认未包含 PDB、日志或截图资产。

详细结果见：

- [v1.0.3 E2E 测试报告](./e2e-test-report-v1.0.3.md)
- [v1.0.3 发布证据与日志索引](./release-evidence-v1.0.3.md)
