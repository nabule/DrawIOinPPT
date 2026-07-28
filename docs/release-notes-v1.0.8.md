# v1.0.8

[中文](./release-notes-v1.0.8.md) | [English](./release-notes-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.8` 将 Word 复杂图形交互改为“不卡优先”。完整 Draw.io XML 继续保存在 `.docx` 内，正常图片 `AlternativeText` 只保存轻量引用；同时 Word 完全取消 `WindowSelectionChange` 订阅和启动时选区读取。选中、拖动、缩放和重新定位期间不再执行插件元数据识别或功能区刷新，用户先选中图片，再点击“重新编辑”等命令。

## 主要变化

- Word 正常写入时，`Document.CustomXMLParts` 保存完整 envelope 和压缩后的 `DrawioXml`，仍是文档内主存储。
- `InlineShape/Shape.AlternativeText` 改为不含 `DrawioXml` 的轻量 envelope，只保留 `formatVersion`、`diagramId`、图形名称、编辑模式、编辑目标、sidecar 路径和更新时间。
- Word 不再监听普通选区变化，也不在启动时读取选区、图片元数据或扫描文档孤儿部件。
- “重新编辑 / 刷新 / 绑定 / 清除绑定”保持可点击，并在点击时读取和验证实时选区；功能区使用固定的“选中后点击操作”提示。
- Word 移除“自动打开”入口；PowerPoint 的选区事件和 `AutoOpenOnSelection` 行为不变。双击 Word 图片仍属于显式编辑动作。
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
- `word-selection-event-e2e.ps1` 通过：发布 DLL 报告 `NoPassiveSelectionMetadataPath=True`，显式读取仍能识别受管和普通浮动图片。
- `word-url-e2e.ps1` 通过：普通 URL 创建、重开和编辑链路从 `Document.CustomXMLParts` 验证完整数据。
- `word-url-addin-host-e2e-safety-test.ps1` 通过：实际宿主脚本的注册、设置、进程和清理安全约束满足断言。
- `word-url-addin-host-e2e.ps1` 通过：真实 `WINWORD.EXE` 中先选中图片，再调用显式“重新编辑”入口；`WordAddInConnect=True`、`ExplicitEditAutomationAvailable=True`、`ExplicitEditCommandInvoked=True`、`ActualWordUrlEditorSaved=True`。
- v1.0.8 临时安装包完整 Office E2E 为 9/9 通过，其中新增 `InstalledWordComplexMetadataE2E`，并覆盖发布打包、安装、Office 注册/加载、URL smoke、真实 Word URL 宿主、PowerPoint URL 回写和桌面导出。
- v1.0.8 发布包已生成并包含 `scripts\word-complex-metadata-e2e.ps1`、中英文报告和证据索引。完整结果见 [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md) 和 [v1.0.8 发布证据](./release-evidence-v1.0.8.md)。
- full E2E 使用 PID 和启动时间精确清理测试拥有的 Word；构建、打包、安装、编译、临时包卸载、仓库加载项重注册或设置恢复失败都进入统一 catch/finally、FatalError 报告和尾部 `exit 1`，清理与统一出口安全约束测试已通过。
- 发布包显式补齐架构、回归清单、优化 backlog 和 v1.0.5 历史 E2E 中英文文档，压缩前包内 Markdown 链接检查为 `0` 个缺失。
- v1.0.8 已安装到标准当前用户目录 `%LOCALAPPDATA%\Greensoft\DrawioPpt`；`PACKAGE.txt` 为 `v1.0.8 / Release / x64`，三项 DLL 与发布包 SHA256 一致，Word / PowerPoint `CodeBase` 指向该标准目录且加载验证通过。
- 安装后的真实 Word UI 线程压力对照使用 120 个 SVG 元素和 254,606 字符 Draw.io XML，执行两轮交叉顺序测试，每张图片每轮预热 4 秒后测量 12 秒，并分别记录选择、位置和尺寸延迟。普通/受管合并 P95 为 7.501 / 5.833 ms、388.910 / 368.568 ms、112.280 / 103.247 ms，受管/普通比例为 0.778 / 0.948 / 0.920。比较、绝对延迟、最小样本、最终失败和无响应门槛全部通过；测量实例加载项连接，按目标 Shape 分类确认 135 次事件且缺失为 0，安装态运行时反射输出 `NoPassiveSelectionMetadataPath=True`。本样本只支持“未观察到超过门槛的额外差异且绝对延迟未超过本次门槛”，不证明绝对延迟的具体来源。
- 压力对照保存重开后，受管图形位置和尺寸保持，图片轻量 `AlternativeText` 为 382 字符，`Document.CustomXMLParts` 可恢复 254,606 字符完整源 XML；真实 URL 宿主回归同时验证 `ActualWordUrlEditorSaved=True`。完整命令和路径规范化证据快照见 [Word UI 线程压力验收报告](./word-ui-thread-stress-report-v1.0.8.md)。

v1.0.8 仍处于待公开发布状态。标准本地安装和自动化 Word UI 线程压力对照已经完成；当前 Codex 交互桌面因系统拒绝光标访问（Win32 error 5），无法代替用户执行真实指针连续拖动，因此普通/受管图形各连续 10 秒的人工指针拖动、缩放、重定位和“重新编辑”按钮点击仍保留为待确认项，本说明不把自动化压力结果冒充人工鼠标验收。
