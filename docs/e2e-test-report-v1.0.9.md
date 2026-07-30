# v1.0.9 Office E2E 测试报告

本报告记录 v1.0.9 发布前的完整 Office E2E。完整运行会临时安装发布包，验证 Word、PowerPoint 加载、编辑与导出流程，并在安装验证阶段直接比对两个宿主返回的 BuildId 与 `BuildInfo.txt`。

## 结果

`13/13 PASS`，且 `FullE2ECleanupSucceeded=True`。

| 检查 | 结果 |
| --- | --- |
| 发布包与 Markdown 链接校验 | 通过 |
| 旧安装目录内子目录升级 | 通过 |
| 临时安装与 Office 注册 | 通过 |
| Word、PowerPoint 实际 COM 加载和版本回调 | 通过 |
| URL 编辑器、Word 实际宿主回写、复杂元数据 | 通过 |
| Word/PowerPoint SVG 等比处理、预览、URL 回写 | 通过 |
| draw.io Desktop 导出 | 通过 |

生成报告：`artifacts\test-reports\full-e2e-v1.0.9.md`。
