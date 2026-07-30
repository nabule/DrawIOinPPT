# v1.0.10 Office E2E 测试报告

本报告记录 v1.0.10 发布前的完整 Office E2E。它会临时安装发布包，验证 Word、PowerPoint 加载、编辑与导出流程，以及安装包脚本的 ASCII 编码和失败退出码处理。

## 结果

`14/14 PASS`，且 `FullE2ECleanupSucceeded=True`。

| 检查 | 结果 |
| --- | --- |
| 发布包、Markdown 链接和 ZIP 解压哈希校验 | 通过 |
| 发布包 PowerShell ASCII 与 `install.cmd` 失败退出码检查 | 通过 |
| 旧安装目录内子目录升级 | 通过 |
| 临时安装、Office 注册和 Word/PowerPoint 实际 COM 版本回调 | 通过 |
| URL 编辑器、Word 实际宿主回写、复杂元数据 | 通过 |
| Word/PowerPoint SVG 等比处理、Word 预览和 PowerPoint URL 回写 | 通过 |
| draw.io Desktop 导出 | 通过 |

生成报告：`artifacts\test-reports\full-e2e-v1.0.10.md`。
