# v1.0.8 Word UI 线程压力测试报告

[中文](./word-ui-thread-stress-report-v1.0.8.md) | [English](./word-ui-thread-stress-report-v1.0.8.en.md) | [脱敏机器可读证据](./evidence/word-ui-thread-stress-v1.0.8.json)

## 目的与边界

本测试在标准本地安装的真实 Word 中，把同一份用户复杂 Draw.io 源图渲染为宽度 `620 px`、保持源宽高比的 PNG，再创建视觉、尺寸、锚点和环绕语义一致的普通图片与受管图片。两者均为浮动 `Shape`，环绕方式为 `Front`。脚本在可见 Word STA UI 线程上执行相同的选中、位置和尺寸属性更新，用于比较受管元数据是否带来可观测差异。

本次测量使用的标准本地 Word DLL SHA256 为 `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D`，与源码 Release 和最终发布包一致。

测试通过 Word COM 更新 `Shape` 属性，不注入真实鼠标指针。`PointerInputCovered=false`；另行尝试 Windows 指针接口时，操作系统返回 Win32 error 5（拒绝访问）。因此本报告不是鼠标拖动验收，也绝不把自动化结果表述为“真实鼠标通过”。

Word 加载项不订阅 `WindowSelectionChange`。测试用独立计数器验证 Word 实际发出的目标图片选区事件；用户在正常使用时必须先选中图片，再点击“重新编辑”等命令，插件才读取实时选区并启动编辑，单纯选中、拖动或缩放不会自动打开编辑器。

## 可复现命令

测试前关闭所有 Word 窗口；`<USER_DRAWIO_SOURCE>` 仅表示调用者提供的本地源文件，不在报告或证据中保存绝对附件路径：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 `
  -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" `
  -DrawioSourcePath "<USER_DRAWIO_SOURCE>" `
  -PictureWrapMode Front `
  -PictureRenderFormat Png `
  -PreviewPixelWidth 620 `
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
  -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json"
```

脚本在包含 150 个正文段落、19,456 个正文字符、8 页、2 个表格和 3 张辅助图片的文档中执行两轮交叉顺序测试：第一轮“普通→受管”，第二轮“受管→普通”。源图含 66 个 `mxCell`，PNG 输出为 `620×876 px`，源图和显示图宽高比均约为 `0.707763`，没有边框像素，普通图与受管图保持相同尺寸、锚点和 `Front` 环绕。

## 结果

执行时间：2026-07-29。

| 指标 | 普通图片 | 受管图片 |
| --- | ---: | ---: |
| 两轮合计时长 | 20.895 s | 20.624 s |
| 选择操作数 | 49 | 49 |
| 选择 P95 | 37.136 ms | 49.044 ms |
| 位置更新数 | 25 | 25 |
| 位置更新平均延迟 | 363.771 ms | 344.120 ms |
| 位置更新 P95 | 524.962 ms | 488.881 ms |
| 位置更新最终失败 | 0 | 0 |
| 尺寸更新数 | 24 | 24 |
| 尺寸更新 P95 | 83.134 ms | 71.072 ms |
| 尺寸更新最终失败 | 0 | 0 |
| Word 无响应采样 | 0 | 0 |

- 两轮合并后的受管/普通位置 P95 比例为 `0.931`，差值为 `-36.081 ms`；受管图位置更新与尺寸更新均未比普通图更慢。
- 第一轮普通/受管位置 P95 为 `540.403/499.107 ms`，差值 `-41.296 ms`；第二轮为 `522.370/488.881 ms`，差值 `-33.489 ms`。
- 目标图片选择操作共 `98` 次，独立计数器确认 `98` 次，缺失 `0`，即 `98/98`，`SelectionChangeEventsPassed=true`。
- 保存并重开后宽高比、位置、尺寸和 `Document.CustomXMLParts` 中的 26,106 字符源 XML 均保留；`WordAddInConnect=true`，清理后没有残留 `WINWORD`。
- 绝对位置 P95 门槛为 `300 ms`，普通图和受管图均未通过；另一个 `17×24 px` 纯色 PNG 基线也未通过同一绝对门槛。因此不能把绝对延迟归因于 Draw.io 图形复杂度或受管元数据，也不能宣称自动化绝对性能验收通过。
- 机器结果为 `ComparisonPassed=false`、`AbsoluteLatencyGatePassed=false`、`OverallPassed=false`。这与完整 Office E2E 的功能性 `12/12 PASS` 是两类门槛，不能混写。

## 结论

最终 620px PNG 样本记录了普通/受管位置 P95 `524.962/488.881 ms`、差 `-36.081 ms`，并确认 `98/98` 目标选区事件；受管图没有增加位置更新 P95。由于绝对 `300 ms` 门槛失败，且 `PointerInputCovered=false`、Windows 指针接口被拒绝，本报告仍不作“鼠标拖动通过”结论。真实鼠标的连续拖动、缩放、重定位和点击“重新编辑”仍是独立人工验收项。

完整的发布用机器可读摘要见[脱敏证据](./evidence/word-ui-thread-stress-v1.0.8.json)。其中用户名、绝对附件路径、临时运行 GUID 和进程身份已移除；保留了源报告的关键配置、统计量、门槛结果和指针覆盖边界。
