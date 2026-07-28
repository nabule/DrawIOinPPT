# v1.0.8

[中文](./release-notes-v1.0.8.md) | [English](./release-notes-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.8` 优化 Word 复杂 Draw.io 图形的图片元数据读取路径。完整 Draw.io XML 继续保存在 `.docx` 内，但不再在正常图片 `AlternativeText` 中重复保存，减少 `WindowSelectionChange` 同步期间的大 XML 反序列化，从而改善复杂图形选中、拖动和缩放时的响应。

## 主要变化

- Word 正常写入时，`Document.CustomXMLParts` 保存完整 envelope 和压缩后的 `DrawioXml`，仍是文档内主存储。
- `InlineShape/Shape.AlternativeText` 改为不含 `DrawioXml` 的轻量 envelope，只保留 `formatVersion`、`diagramId`、图形名称、编辑模式、编辑目标、sidecar 路径和更新时间。
- 如果 `CustomXMLParts.Upsert` 失败，图片 `AlternativeText` 会保留全量 envelope，避免保存过程中丢失源数据。
- 旧版只在图片 `AlternativeText` 中保存全量 envelope 的文档，会在首次读取时自动迁入 `CustomXMLParts`；迁移成功后图片改为轻量引用，重复读取不会创建新的 XML part。
- Word 普通 URL 回归和实际加载项宿主回归改为严格验证文档级主存储，避免把轻量引用误判为完整回写。
- PowerPoint 的 `Presentation.CustomXMLParts + Shape.Tags + Shape.AlternativeText` 存储策略没有随本次 Word 优化改变。

## 兼容取舍

Office 不保证单独复制一张 Word 图片到另一文档时，同时复制原文档的 `CustomXMLParts`。正常轻量 `AlternativeText` 不含 Draw.io XML，因此目标文档中可能无法恢复编辑源。

建议跨文档交付时复制整份 `.docx`，或事先保存/保留 sidecar `.drawio` 或 Draw.io XML。只有旧版图片或主存储写入失败的回退图片，`AlternativeText` 才可能仍含全量数据；不能把它当作所有图片都具备的恢复来源。

## 当前已通过的验证

- Release x64 解决方案构建通过：`0` 个警告，`0` 个错误。
- `word-complex-metadata-e2e.ps1` 通过：大 XML 主存储、轻量引用、保存关闭重开、失败回退、旧版迁移、二次读取 part ID 稳定及无残留 Word 进程均满足断言。
- `word-selection-event-e2e.ps1` 通过：真实 Word 选区识别与无轮询约束满足断言。
- `word-url-e2e.ps1` 通过：普通 URL 创建、重开和编辑链路从 `Document.CustomXMLParts` 验证完整数据。
- `word-url-addin-host-e2e-safety-test.ps1` 通过：实际宿主脚本的注册、设置、进程和清理安全约束满足断言。
- `word-url-addin-host-e2e.ps1` 通过：真实 `WINWORD.EXE` 中 `WordAddInConnect=True`、`ActualWordUrlEditorSaved=True`。
- v1.0.8 临时安装包完整 Office E2E 为 9/9 通过，其中新增 `InstalledWordComplexMetadataE2E`，并覆盖发布打包、安装、Office 注册/加载、URL smoke、真实 Word URL 宿主、PowerPoint URL 回写和桌面导出。
- v1.0.8 发布包已生成并包含 `scripts\word-complex-metadata-e2e.ps1`、中英文报告和证据索引。完整结果见 [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md) 和 [v1.0.8 发布证据](./release-evidence-v1.0.8.md)。
- full E2E 使用 PID 和启动时间精确清理测试拥有的 Word；构建、打包、安装、编译、临时包卸载、仓库加载项重注册或设置恢复失败都进入统一 catch/finally、FatalError 报告和尾部 `exit 1`，清理与统一出口安全约束测试已通过。
- 发布包显式补齐架构、回归清单、优化 backlog 和 v1.0.5 历史 E2E 中英文文档，压缩前包内 Markdown 链接检查为 `0` 个缺失。

v1.0.8 仍处于待公开发布状态。标准本地目录的最终安装和人工 10 秒拖动、缩放、重定位及重开编辑对比属于后续验收，本说明不提前宣称这些项目已经通过。
