# 回归清单

[中文](./regression-checklist.md) | [English](./regression-checklist.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## v1.0.8 本轮执行状态

| 状态 | 项目 |
| --- | --- |
| 通过 | 临时安装发布包完整 Office E2E 12/12；该结果仅代表隔离临时安装。 |
| 通过 | 标准本地目录 `%LOCALAPPDATA%\Greensoft\DrawioPpt` 安装、DLL 哈希、Office 注册和真实 COM 加载。 |
| 未通过绝对门槛 | 最终 620px PNG / `Front` 富文档压力样本中，普通/受管位置 P95 为 556.177 / 574.474 ms，差 18.297 ms，94/94 目标选区事件确认；两者均未通过绝对 300 ms 门槛，17×24 纯色 PNG 基线也未通过。 |
| 待人工确认 | `PointerInputCovered=false`；Windows `GetCursorPos` 访问被拒绝，`SetIsBorderRequired` 接口不受支持，自动化不能代替用户执行真实指针拖动、缩放、重定位和命令点击。 |

## 1. 构建与注册

- `scripts/dev-check.ps1` 通过
- `scripts/build.ps1` 通过
- `scripts/register-office-addins.ps1` 可同时完成 PowerPoint 和 Word 注册
- PowerPoint 中 `Greensoft.DrawioPptAddIn` 可正常加载

## 2. Ribbon 与选择状态

- Ribbon 可正常显示
- Ribbon 分组为“创建 / 当前图形 / 工作流 / 信息”
- Word 状态固定显示“对象：选中后点击操作 / 状态：点击按钮时识别”，选择变化不会触发功能区刷新
- Word 的“重新编辑 / 刷新 / 绑定 / 清除绑定”保持可点击，并在点击时验证实时选区
- Word 不显示“自动打开”开关；PowerPoint 的 `AutoOpenOnSelection` 行为不变
- 按钮图标可正常显示，不出现空白占位

## 3. 桌面模式

- 可自动探测本地 `draw.io` 路径
- 可新建图形
- 可打开桌面编辑器
- 未保存的空白 PPT 中新建图形时，sidecar 会落到绝对临时路径，不会生成相对 `drawio\*.drawio`
- 已保存或另存为后的 PPT，再次编辑时 sidecar 会自动切换到当前 PPT 目录下
- 将 PPT 拷贝到另一台机器后，再次编辑时 sidecar 会按当前机器上的 PPT 路径重建
- 保存 `.drawio` 后 PPT 中 SVG 自动刷新
- 手动刷新按钮可用

## 4. URL 模式

- 设置窗口 `Test URL` 可通过
- 可打开已有图形
- `init -> load -> save -> export` 可走通
- SVG 与 XML 可回写到 PPT
- 新建图形后保存 PPT、关闭并重新打开，仍可再次编辑并回写
- 日志中可看到 URL 模式关键事件
- `scripts\word-url-addin-host-e2e.ps1` 在真实 `WINWORD.EXE` 中通过：加载项已连接、选中受管图片后显式调用“重新编辑”并完成回写、`ExplicitEditCommandInvoked=True`、`%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2` 存在，新增日志没有 `E_ACCESSDENIED`。

## 5. 存储与迁移

- 新图形会写入 `Shape.Tags`
- 新图形会写入 `Presentation.CustomXMLParts`
- 重新打开 PPT 后可从 `CustomXMLParts` 恢复源 XML 并继续编辑
- 旧版只含 `AlternativeText` 的图形会自动迁移
- 删除图形绑定后，不再保留无引用的孤儿 `CustomXMLPart`
- 生成出的 SVG 包含 draw.io `content` 元数据

## 6. 展示对象替换保真

- 替换后位置不跳动
- 替换后尺寸不跳动
- 替换后旋转保留
- 替换后名称保留
- 替换后超链接保留
- `scripts\word-preview-image-provider-e2e.ps1` 覆盖正常 620px 等比例 PNG、导出失败时 620×310 等比例轻量 PNG 占位预览，以及临时 XML 重试、延迟和启动清理；不得回退复杂 SVG。
- `scripts\word-svg-aspect-ratio-e2e.ps1` 在真实 Word 中通过：2:1 PNG 新建和替换保持 2:1，均为 `wdWrapFront` 浮动图片；`VerticalRatio=0.25`、`VerticalContained=True`；`FailedReplacementPreservedOriginal=True` 验证替换先建新图、成功后才删原图。
- `scripts\powerpoint-svg-aspect-ratio-e2e.ps1` 在真实 PowerPoint 中通过：直接插入原始 SVG，按边界 `contain` 等比居中，替换后保持源比例且不扩展 `viewBox` 产生图内留白。

## 7. 异常与日志

- 桌面路径缺失时有明确提示
- URL 初始化超时有明确提示
- URL 导出超时有明确提示
- 宿主入口异常会落日志并弹出错误框
- 日志文件可持续追加写入

## 8. 注册与卸载

- 卸载脚本可正常移除注册项
- 卸载后 PowerPoint 和 Word 不再加载对应插件

## 9. Word 选区与图片操作

- `scripts\word-selection-sync-test.ps1` 通过：`SelectionMonitor` 不含 `WindowSelectionChange` / `SelectionChanged`，`AddInHost` 不含选区 Timer、被动选区处理器或启动时选区读取。
- `scripts\word-selection-event-e2e.ps1` 在真实 Word COM 中通过：发布 DLL 输出 `NoPassiveSelectionMetadataPath=True`，并验证显式读取能够识别受管浮动 `Shape` 和普通图片。
- 执行 Word E2E 前关闭 Word，且不要与其他 Word 自动化测试并行运行；脚本会还原插件设置并仅清理本次测试创建的自动化 Word 进程。
- 普通选中、移动、缩放和重新定位期间不得读取图片 `AlternativeText`、`CustomXMLParts` 或刷新功能区；先选中图片，再点击“重新编辑”“刷新”“绑定”或“清除绑定”，命令才读取当前选区。

## 10. Word 复杂元数据与拖动性能

自动化检查：

- `scripts\word-complex-metadata-e2e.ps1` 使用大 Draw.io XML 验证：完整正文保存在 `Document.CustomXMLParts`，图片 `AlternativeText` 不含 `DrawioXml` 且长度不超过 2048 字符。
- 保存、关闭并重开 `.docx` 后，文档级全量 envelope 与图片轻量引用都仍然存在，引用字段保持一致。
- 模拟 `CustomXMLParts.Upsert` 失败时，图片 `AlternativeText` 保留全量 envelope；保存、关闭、重开后回退数据仍可读取。
- 旧版仅在 `AlternativeText` 保存全量 envelope 的图片首次读取时自动迁入 `CustomXMLParts`；二次读取不新增部件，相关 XML part ID 保持稳定。
- `scripts\word-url-e2e.ps1` 严格从 `Document.CustomXMLParts` 验证 URL 模式的创建、重开、编辑和最终回写，不以轻量 `AlternativeText` 冒充完整主存储。
- `scripts\word-url-addin-host-e2e.ps1` 严格从文档级主存储验证实际宿主回写，`WordAddInConnect=True` 且 `ActualWordUrlEditorSaved=True`。
- 每个 Word 自动化脚本结束后检查没有残留 `WINWORD.EXE`；发现残留进程即判定失败。
- 安装态可见 Word UI 线程压力对照使用用户复杂源图生成 620×876 PNG，普通/受管图均为等比例 `Front` 浮动图片。普通/受管位置 P95 为 556.177 / 574.474 ms，差 18.297 ms；94/94 目标 Shape 选区事件确认，缺失为 0。绝对 300 ms 门槛未通过，且 17×24 纯色 PNG 基线同样失败；`ComparisonPassed=false`、`AbsoluteLatencyGatePassed=false`、`OverallPassed=false`。安装态运行时反射仍独立确认 `NoPassiveSelectionMetadataPath=True`，但该源码/反射事实不能冒充性能或鼠标验收通过。完整结果见 [v1.0.8 Word UI 线程压力测试报告](./word-ui-thread-stress-report-v1.0.8.md)。

发布前人工验收：

| 状态 | 检查项 |
| --- | --- |
| 自动化未覆盖，待指针确认 | `PointerInputCovered=false`；Windows 指针接口被拒绝/不支持，且绝对 300 ms 门槛失败。仍需用户用真实指针对每个对象各连续 10 秒拖动、缩放和重新定位，确认视觉跟随。 |
| 自动化显式命令通过，待人工按钮确认 | 保存重开后位置、尺寸及 26,106 字符源 XML 已恢复，真实 URL 宿主 E2E 已通过“选中后显式重新编辑”入口完成回写；这不等于用户真实点击已验收，仍需本机人工点击一次“重新编辑”。 |
