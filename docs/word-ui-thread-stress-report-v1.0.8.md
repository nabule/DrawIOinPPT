# v1.0.8 Word UI 线程压力验收报告

[中文](./word-ui-thread-stress-report-v1.0.8.md) | [English](./word-ui-thread-stress-report-v1.0.8.en.md) | [路径规范化证据快照](./evidence/word-ui-thread-stress-v1.0.8.json)

## 目的与边界

本测试在标准本地安装的真实 Word 中，对相同视觉和尺寸的普通 SVG 图片与受管 Draw.io 图片执行同位置、同操作序列的 UI 线程压力对照。每个样本均先选择正文中性位置，等待 75 ms 消息处理，再分别计时目标图片选择事件和位置/尺寸更新；测量实例必须已连接加载项。测试自身的独立事件计数器按 Shape 名称分类，预热后清零对应对象计数，确认 Word 对目标图片实际发出 `WindowSelectionChange`，正文选择不能充入确认数；发布 DLL 的运行时反射另行确认 `NoPassiveSelectionMetadataPath=True`，加载项本身不订阅该事件。

测试通过 Word COM 在可见 Word STA UI 线程更新 `Shape` 属性，不注入真实鼠标指针。因此，它不是人工指针拖动验收，不能证明图形一定逐像素跟随真实指针。

## 可复现命令

测试前关闭所有 Word 窗口：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 `
  -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" `
  -DurationSeconds 12 `
  -MaximumP95Ratio 1.25 `
  -MaximumPerRoundP95Ratio 1.50 `
  -MaximumSelectionP95DeltaMs 100 `
  -MaximumAbsoluteSelectionP95Ms 2000 `
  -MaximumAbsoluteP95Ms 2000 `
  -MaximumAbsoluteResizeP95Ms 750 `
  -MinimumOperationsPerRound 10 `
  -WarmupSeconds 4 `
  -ResizeOperationsPerRound 12 `
  -SelectionSettleMilliseconds 75 `
  -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-v1.0.8.json"
```

脚本生成 120 个 SVG 元素和 1,200 个 Draw.io 单元。它执行两轮交叉顺序测试：第一轮“普通→受管”，第二轮“受管→普通”；每张图片每轮先执行 4 秒独立预热，再连续执行至少 12 秒选择事件与位置更新，并执行 12 次单轴尺寸更新。普通和受管图片在各自测量时单独显示于同一页面位置，降低视口、首次布局、重叠渲染和固定顺序偏差。

## 结果

执行时间：2026-07-28。

| 指标 | 普通图片 | 受管图片 |
| --- | ---: | ---: |
| 两轮合计时长 | 24.184 s | 24.516 s |
| 选择事件计时次数 | 66 | 69 |
| 选择事件平均延迟 | 3.594 ms | 3.368 ms |
| 选择事件 P95 | 7.501 ms | 5.833 ms |
| 位置更新次数 | 42 | 45 |
| 位置更新平均延迟 | 311.822 ms | 287.630 ms |
| 位置更新 P95 | 388.910 ms | 368.568 ms |
| 位置更新最终失败 | 0 | 0 |
| 单轴尺寸更新次数 | 24 | 24 |
| 单轴尺寸更新 P95 | 112.280 ms | 103.247 ms |
| 尺寸更新最终失败 | 0 | 0 |
| Word 无响应采样 | 0 | 0 |

- 第一轮普通/受管位置 P95 为 `387.316/368.568 ms`，比例 `0.952`；第二轮为 `388.910/314.240 ms`，比例 `0.808`。
- 两轮合并后的受管/普通选择、位置和尺寸 P95 比例分别为 `0.778`、`0.948` 和 `0.920`，均未超过 `1.25` 门槛。单轮最高选择、位置和尺寸比例分别为 `1.015`、`0.952` 和 `1.151`，均未超过 `1.50`；合并选择额外延迟为 `-1.668 ms`。
- 四组选择、位置和尺寸 P95 最高分别为 `8.766 ms`、`388.910 ms` 和 `129.184 ms`，未超过 `2000 / 2000 / 750 ms` 绝对门槛；每组位置样本至少 `21` 个，超过最低 `10` 个要求。
- 测量实例 `MeasurementWordAddInConnect=True`；测试计数器共观察到 `414` 次事件，按目标 Shape 分类确认 `135` 次，对应 `135` 次目标选择操作，缺失数为 `0`，`SelectionChangeEventsPassed=True`。安装态运行时反射同时输出 `NoPassiveSelectionMetadataPath=True`。
- Draw.io XML：`254,606` 字符；受管图片轻量 `AlternativeText`：`382` 字符。
- 保存重开后位置和尺寸保持，`Document.CustomXMLParts` 恢复 `254,606` 字符完整源 XML。
- `SelectionComparisonPassed=True`、`ComparisonPassed=True`、`AbsoluteLatencyGatePassed=True`、`MinimumSamplesPassed=True`、`WordAddInConnect=True`、`CleanupResidualWinWord=False`、`OverallPassed=True`。

本次样本只支持“未观察到图片元数据路径带来超过门槛的额外差异，且绝对延迟未超过本次验收门槛”。它不证明绝对延迟的具体来源，也不替代对普通和受管图片各自执行人工连续 10 秒指针拖动、缩放、重定位以及“重新编辑”按钮点击。

完整机器可读结果见[路径规范化证据快照](./evidence/word-ui-thread-stress-v1.0.8.json)。该快照只把真实输出中的用户目录替换为 `%LOCALAPPDATA%` 和 `%TEMP%`，其余字段与当次脚本输出一致。
