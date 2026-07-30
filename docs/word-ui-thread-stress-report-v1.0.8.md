# v1.0.8 Word UI 线程压力测试报告

[中文](./word-ui-thread-stress-report-v1.0.8.md) | [English](./word-ui-thread-stress-report-v1.0.8.en.md) | [脱敏机器可读证据](./evidence/word-ui-thread-stress-v1.0.8.json)

## 目的与边界

本测试在标准本地安装的真实 Word 中，把同一份用户复杂 Draw.io 源图渲染为宽度 `1240 px`、保持源宽高比的 PNG，再创建视觉、尺寸、锚点和环绕语义一致的普通图片与受管图片。两者均为浮动 `Shape`，环绕方式为 `Front`。脚本在可见 Word STA UI 线程上执行相同的选中、位置和尺寸属性更新，用于比较受管元数据是否带来可观测差异。

本次测量使用的标准本地 Word DLL SHA256 为 `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8`，与源码 Release 和最终发布包一致。

测试通过 Word COM 更新 `Shape` 属性，不注入真实鼠标指针。`PointerInputCovered=false`。此前尝试真实桌面接口时分别得到 `GetCursorPos=AccessDenied` 和 `SetIsBorderRequired=Unsupported`。因此本报告不是鼠标拖动验收，也绝不把自动化结果表述为“真实鼠标通过”。用户已决定略过独立人工指针验收，状态为“略过/未验收”。

Word 加载项不订阅 `WindowSelectionChange`。测试用独立计数器验证 Word 实际发出的目标图片选区事件；用户在正常使用时必须先选中图片，再点击“重新编辑”等命令，插件才读取实时选区并启动编辑，单纯选中、拖动或缩放不会自动打开编辑器。

## 可复现命令

测试前关闭所有 Word 窗口；`<USER_DRAWIO_SOURCE>` 仅表示调用者提供的本地源文件，不在报告或证据中保存绝对附件路径：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 `
  -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" `
  -DrawioSourcePath "<USER_DRAWIO_SOURCE>" `
  -PictureWrapMode Front `
  -PictureRenderFormat Png `
  -PreviewPixelWidth 1240 `
  -DurationSeconds 10 `
  -MaximumP95Ratio 1.25 `
  -MaximumPerRoundP95Ratio 1.25 `
  -MaximumP95DeltaMs 50 `
  -MaximumSelectionP95DeltaMs 50 `
  -MaximumAbsoluteSelectionP95Ms 300 `
  -MaximumAbsoluteP95Ms 300 `
  -MaximumAbsoluteResizeP95Ms 300 `
  -MinimumOperationsPerRound 10 `
  -WarmupSeconds 4 `
  -ResizeOperationsPerRound 12 `
  -SelectionSettleMilliseconds 75 `
  -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json"
```

脚本在包含 150 个正文段落、19,456 个正文字符、8 页、2 个表格和 3 张辅助图片的文档中执行两轮交叉顺序测试：第一轮“普通→受管”，第二轮“受管→普通”。源图含 66 个 `mxCell`，PNG 输出为 `1240×1753 px`。脚本另行导出 SVG 探针取得独立源比例，再与 PNG 比较，同时验证 draw.io 使用 `--border 0`。透明边扫描要求整行或整列 100% 像素的 Alpha 不高于 8，并保留 2px 栅格化边缘容差；1px 触边线会立即终止该方向的空白判定。不带 Alpha 的 24bpp PNG/JPEG 明确返回 `BlankPaddingDetectionSupported=false` 和空的 `BorderPaddingPassed`，不得伪报成像素检测通过。本次实际 PNG 属于该不透明类型，当前无强制 1:1 和无导出边框的证据来自独立比例误差 `0.000236` 与 `--border 0`。普通图与受管图保持相同尺寸、锚点和 `Front` 环绕。

## 结果

执行时间：2026-07-29。

| 指标 | 普通图片 | 受管图片 |
| --- | ---: | ---: |
| 两轮合计时长 | 21.689 s | 20.608 s |
| 选择操作数 | 48 | 48 |
| 选择 P95 | 68.866 ms | 74.243 ms |
| 位置更新数 | 24 | 24 |
| 位置更新平均延迟 | 202.477 ms | 196.560 ms |
| 位置更新 P95 | 321.844 ms | 325.764 ms |
| 位置更新最终失败 | 0 | 0 |
| 尺寸更新数 | 24 | 24 |
| 尺寸更新 P95 | 89.633 ms | 122.968 ms |
| 尺寸更新最终失败 | 0 | 0 |
| Word 无响应采样 | 0 | 0 |

- 两轮合并后的受管/普通位置 P95 比例为 `1.012`，差值为 `3.920 ms`。
- 选择 P95 比值为 `1.078`、差值 `5.377 ms`；缩放 P95 比值为 `1.372`、差值 `33.335 ms`。由于单轮门槛波动，相对门槛为 `ComparisonPassed=false`、`NoAdditionalMetadataPathDifferenceObserved=false`。
- 目标图片选择操作共 `96` 次，独立计数器确认 `96` 次，缺失 `0`，即 `96/96`，`SelectionChangeEventsPassed=true`。
- 保存并重开后宽高比、位置、尺寸和 `Document.CustomXMLParts` 中的 26,106 字符源 XML 均保留；`WordAddInConnect=true`，清理后没有残留 `WINWORD`。
- 受管图两轮位置 P95 为 `303.892/337.339 ms`，普通对照图为 `328.256/318.748 ms`，均存在超过 `300 ms` 的轮次，导致 `AbsoluteLatencyGatePassed=false`。此前确认轮相对门槛曾通过且受管图两轮低于 300ms，说明时延存在波动，但不能据此放宽门槛。
- 机器结果为 `ComparisonPassed=false`、`AbsoluteLatencyGatePassed=false`、`OverallPassed=false`；两类图片的选择、位置和缩放最终失败均为 0，Word 无响应采样均为 0。这与完整 Office E2E 的功能性 `13/13 PASS` 是两类门槛，不能混写。

## 结论

最终 1240px PNG 样本独立验证了源比例和 `--border 0`；实际不透明 PNG 的透明留白像素检测明确标记为不支持。最新轮普通/受管选择 P95 为 `68.866/74.243 ms`、位置 P95 为 `321.844/325.764 ms`、缩放 P95 为 `89.633/122.968 ms`，并确认 `96/96` 目标选区事件。相对与绝对门槛均未通过，`OverallPassed=false`。`PointerInputCovered=false`，用户已决定略过真实鼠标人工验收；本报告不作“鼠标拖动通过”结论。

完整的发布用机器可读摘要见[脱敏证据](./evidence/word-ui-thread-stress-v1.0.8.json)。其中用户名、绝对附件路径、临时运行 GUID 和进程身份已移除；保留了源报告的关键配置、统计量、门槛结果和指针覆盖边界。
