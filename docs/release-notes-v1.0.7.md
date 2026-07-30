# v1.0.7

[中文](./release-notes-v1.0.7.md) | [English](./release-notes-v1.0.7.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.7` 修复 Word 使用 URL 编辑器时可能出现的“无法初始化 URL 编辑器，拒绝访问”。本次不改变 Draw.io XML、SVG、图片布局或 Office 文档内的元数据存储格式。

## 这版先解决什么问题

- Word 中使用 URL 编辑器时，部分机器会因为 Office 安装目录不可写而报“拒绝访问”。
- 用户不应该为了打开 WebView2 编辑器去修改 `Program Files` 或 Office 安装目录权限。
- URL 模式失败时需要有真实 Word 宿主回归和日志证据，方便判断是否仍是初始化问题。

## 主要变化

- URL 编辑器现在显式创建 WebView2 环境，并将 profile 放到 `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`。
- Word 和 PowerPoint 不再依赖 WebView2 在 `WINWORD.EXE` 或 `POWERPNT.EXE` 所在目录创建默认 profile，避免 Office 安装目录无写权限时的 `E_ACCESSDENIED`。
- 新增 `webview2-user-data-folder-test.ps1` 源码约束，防止代码回退到宿主默认 profile。
- 新增 `word-url-addin-host-e2e.ps1`：在真实 `WINWORD.EXE` 中临时注册被测 DLL、选择受管图片触发 URL 编辑器，并验证 URL 回写、用户级 profile 与日志。
- 回归脚本只清理由自身 helper 报告且启动时间匹配的 Word 进程；Word 注册、设置和临时目录恢复互相隔离，任一恢复失败都会汇总并令测试失败。
- 完整 Office E2E 现在执行已安装发布 DLL 的真实 Word URL 宿主回归；发布脚本同时强制包含本版本的中英文发布资料。

## 使用提示

安装后请完全退出并重新打开 Word 或 PowerPoint。URL 模式使用当前用户的 LocalAppData 目录，不需要修改 `C:\Program Files\Microsoft Office` 权限。若仍无法初始化，请确认 WebView2 Runtime 已安装，并查看 `%APPDATA%\Greensoft\DrawioPpt\Logs\drawioppt.log` 中的 `UrlEditor` 记录。

完整验收见 [v1.0.7 E2E 测试报告](./e2e-test-report-v1.0.7.md) 和 [发布证据与日志索引](./release-evidence-v1.0.7.md)。
