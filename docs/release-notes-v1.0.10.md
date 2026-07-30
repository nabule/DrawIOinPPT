# DrawioPpt v1.0.10 发布说明

[中文首页](../README.md) | [English](./release-notes-v1.0.10.en.md)

## 这版先解决什么问题

部分 Windows 机器运行 v1.0.9 的 `install.cmd` 时，Windows PowerShell 会把安装脚本中的中文文本按错误编码读取，直接报乱码和语法错误。安装在开始复制文件前就中断，已安装机器无法升级。

## 这版如何解决

- 发布包内的所有 PowerShell 脚本改为 ASCII，避免 Windows PowerShell 5.1 的 UTF-8 无 BOM 解码问题。
- `install.cmd` 现在会保留 PowerShell 原始错误、打印安装失败退出码，并提示先关闭 Word、PowerPoint 后重试。
- 新增发布包编码安全检查，验证每个随包 `.ps1` 以及生成的 `install.cmd` 都满足上述要求。

## 最终如何使用

1. 下载并解压 `DrawioPpt-v1.0.10.zip`。
2. 关闭 Word 和 PowerPoint。
3. 在解压目录双击 `install.cmd`。
4. 成功后打开 Word 或 PowerPoint，在 `Draw.io` 功能区的信息区确认 BuildId 为 `v1.0.10+<git-short-hash>`。

如果安装失败，不要关闭命令窗口。先查看 PowerShell 原始错误和 `DrawioPpt installation failed with exit code ...`，按提示关闭占用程序或处理文件锁定，再运行新包中的 `install.cmd`。不要继续使用会显示乱码的 v1.0.9 安装包。

## 验证

- 发布包的所有 PowerShell 脚本均通过 ASCII 字节检查和 Windows PowerShell 文件解析检查。
- 在 Word 保持打开时，真实 `install.cmd` 保留 PowerShell 原始错误并打印 `DrawioPpt installation failed with exit code 1`。
- 完整 Office E2E `14/14 PASS`，`FullE2ECleanupSucceeded=True`；实际 Word 和 PowerPoint 返回的版本文本均与安装包 `BuildInfo.txt` 一致。

完整证据见 [v1.0.10 E2E 测试报告](./e2e-test-report-v1.0.10.md) 和 [v1.0.10 发布证据](./release-evidence-v1.0.10.md)。
