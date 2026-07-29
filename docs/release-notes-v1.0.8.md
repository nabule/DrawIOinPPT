# DrawioPpt v1.0.8 发布说明

[中文](./release-notes-v1.0.8.md) | [English](./release-notes-v1.0.8.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 主要变化

- Word 选区热路径不再订阅 `WindowSelectionChange`，启动时也不读取选区、图片属性、`AlternativeText` 或 `CustomXMLParts`。普通选中、拖动、缩放和重新定位不执行插件元数据处理。
- “重新编辑”“刷新”“绑定”和“清除绑定”保持可点击；用户先选中图片，再点击命令，插件才读取实时选区并执行相应操作。选中本身不会自动打开编辑器。
- Word 正常展示改为从 Draw.io 源生成宽度 620px、零边框、保持源宽高比的 PNG；图片转换为浮动 `Shape`，环绕固定为 `wdWrapFront`。draw.io Desktop 导出失败时生成 620×310 等比例轻量 PNG 占位预览，不回退复杂 SVG；源 XML 仍在文档元数据中，点击显式命令仍可编辑。
- PowerPoint 继续直接插入原始 SVG：在幻灯片边界内按源比例 `contain` 缩放并居中，不改写 `viewBox`，不为填满边界框制造图内留白。
- Word 重新编辑或刷新时保留图片宽度，按源比例重算高度和位置。完整 Draw.io XML 仍以 `Document.CustomXMLParts` 为主存储，图片 `AlternativeText` 正常只保留轻量引用。
- `CustomXMLParts.Upsert` 失败时，图片回退保留完整 envelope；旧版只在图片 `AlternativeText` 保存全量 envelope 的文档会在首次读取时迁入文档级主存储。

## 兼容取舍

Office 不保证单独复制一张 Word 图片到另一文档时，同时复制源文档的 `CustomXMLParts`。正常轻量 `AlternativeText` 不含 Draw.io XML，因此目标文档可能无法恢复编辑源。跨文档交付时请复制整份 `.docx`，或事先保留 sidecar `.drawio` / Draw.io XML。

Word PNG 负责显示而非主存储。正常路径需要可用的 draw.io Desktop PNG 导出；导出失败时使用轻量 PNG 占位而不是复杂 SVG，不丢失文档级源 XML。

## 当前验证结果

- `Release|x64` 构建通过：0 warnings，0 errors。
- Word 零被动选区路径、WebView2 用户目录、URL 宿主安全约束和 full E2E 清理安全约束均通过。
- 最新 `word-preview-image-provider-e2e.ps1` 覆盖正常 620px 等比例 PNG、失败时 620×310 等比例轻量 PNG 占位预览，以及临时 XML 重试/延迟/启动清理。
- 最新 `word-svg-aspect-ratio-e2e.ps1` 通过：真实 Word 中 2:1 PNG 新建和替换保持 2:1；纵向图 `VerticalRatio=0.25`、`VerticalContained=True`；失败替换 `FailedReplacementPreservedOriginal=True`，证明新图成功前不会删除原图。
- `powerpoint-svg-aspect-ratio-e2e.ps1` 通过：真实 PowerPoint 直接插入原始 SVG，按边界 `contain` 等比居中，替换保持源比例且不扩展 `viewBox`。
- 现有临时安装发布包快照的完整 Office 功能 E2E 为 `12/12 PASS`，覆盖 `InstalledWordSvgAspectE2E`、`InstalledWordPreviewProviderE2E` 和 `InstalledPowerPointSvgAspectE2E`。上述最新 Word 替换/清理增强晚于该包快照，需重打包后重新执行 full E2E。
- full E2E 在执行前快照 Word / PowerPoint 加载项注册树，结束时精确恢复原状态，包括原本不存在的键；最终 `WINWORD` / `POWERPNT` 零残留硬门槛通过，`FullE2ECleanupSucceeded=True`。
- 当前 Core 与 PowerPoint DLL 在源码 Release、发布包和标准本地安装中一致；最新 Word 源码 DLL 为 `F5B68A12…AD70`，而现有发布包/本地安装仍为 `A888A38D…B5D9`，尚未同步。完整 DLL 哈希见[发布证据](./release-evidence-v1.0.8.md)。

## Word 压力与指针边界

最终用户复杂源图样本使用 620×876 PNG、`Front` 浮动布局和富文档基线。普通/受管位置 P95 为 `556.177/574.474 ms`，差值 `18.297 ms`；目标选区事件 `94/94`，缺失 0。

绝对位置 P95 门槛为 `300 ms`，普通图和受管图均未通过；`17×24 px` 纯色 PNG 基线也未通过同一门槛。机器结果为 `ComparisonPassed=false`、`AbsoluteLatencyGatePassed=false`、`OverallPassed=false`，不得写成自动化性能验收通过。

`PointerInputCovered=false`。Windows `GetCursorPos` 访问被拒绝，`SetIsBorderRequired` 接口不受支持，自动化未覆盖真实鼠标输入。因此本发布说明不声称鼠标拖动通过；真实指针拖动、缩放、重定位和人工点击“重新编辑”仍是独立待确认项。

## 发布状态

v1.0.8 仍处于待公开发布状态。当前 ZIP 和标准本地 Word DLL 均早于最新替换/清理增强；公开发布前必须重打包、更新本地安装、重跑 `12/12` full E2E 并复核包内链接。ZIP SHA256 不能写入会再次进入 ZIP 的文档，由发布脚本或包外交付摘要在最终打包后记录。

完整结果见 [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md)、[v1.0.8 发布证据](./release-evidence-v1.0.8.md)和 [Word UI 线程压力测试报告](./word-ui-thread-stress-report-v1.0.8.md)。
