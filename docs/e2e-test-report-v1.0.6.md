# v1.0.6 E2E 测试报告

[中文](./e2e-test-report-v1.0.6.md) | [English](./e2e-test-report-v1.0.6.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 测试环境

- Windows + Microsoft Word Desktop / PowerPoint Desktop
- .NET Framework 4.8，`Release|x64`
- 当前用户级 Office COM Add-in 注册

## 本次回归

| 项目 | 结果 | 关键证据 |
| --- | --- | --- |
| 无轮询红灯 | 通过预期失败 | 实现前约束脚本发现 `_selectionStateTimer` |
| 无轮询绿灯 | 通过 | `WORD_SELECTION_SYNC_TEST_PASS` |
| Release 构建 | 通过 | 0 warnings，0 errors |
| Word 浮动图片 COM E2E | 通过 | `NoSelectionPolling=True`、`ManagedSelectionDetected=True`、`PlainPictureCanBind=True` |
| Word URL 写回 E2E | 通过 | `MockHttpReachable=True`、新建/重开编辑/再次重开均为 `True` |
| 测试环境恢复 | 通过 | URL E2E 输出 `RegisteredWordAddInLoadBehavior=Restored:3`，未遗留 Word 进程 |

## 覆盖说明

`word-selection-event-e2e.ps1` 在真实 Word 中创建受管与普通浮动 `Shape`，验证生产事件读取器的现场选区识别，并反射检查被测 DLL 不含旧的 Timer 字段或 Tick 回调。显式 Word 命令也使用同一现场选区读取路径。

`word-url-e2e.ps1` 覆盖 URL 模式的 `configure -> init -> load -> save -> export`、保存、重新打开和二次编辑。测试 mock 在同一测试进程内运行、先完成 HTTP 自检并记录 iframe 请求；开始前 Word 必须关闭，测试结束恢复已注册加载项的原 `LoadBehavior`。

发布包安装、发布 DLL E2E 与完整 Office E2E 的最终记录见 [发布证据与日志索引](./release-evidence-v1.0.6.md)。
