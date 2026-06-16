# v1.0.4

[中文](./release-notes-v1.0.4.md) | [English](./release-notes-v1.0.4.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.4` 是面向安装脚本收口的发布版本。它补齐仓库开发安装入口和发布包内可见脚本，确保 PowerPoint 与 Word 插件可以通过统一入口同时注册和注销。

## 主要变化

- 新增 `scripts\register-office-addins.ps1`，一次性注册 PowerPoint 与 Word 两个当前用户级 COM Add-in。
- 新增 `scripts\unregister-office-addins.ps1`，一次性注销 PowerPoint 与 Word 两个 COM Add-in。
- 发布包 `install.cmd` / `scripts\install-release.ps1` 改为调用统一注册入口，避免后续安装链路只覆盖 PowerPoint。
- 发布包 `uninstall.cmd` / `scripts\uninstall-release.ps1` 改为调用统一注销入口。
- 打包脚本和 `PACKAGE.txt` 清单加入统一注册/注销脚本。
- README、安装说明、用户手册、架构说明和回归清单均明确双宿主安装方式。

## 安装方式

最终用户从 Release 下载：

```text
DrawioPpt-v1.0.4.zip
```

解压后关闭 PowerPoint 和 Word，再运行：

```powershell
.\install.cmd
```

安装脚本会同时注册：

- `Greensoft.DrawioPptAddIn`
- `Greensoft.DrawioWordAddIn`

仓库开发调试时，先构建，再运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-office-addins.ps1
```

如果只需要排查单个宿主，仍可直接使用 `scripts\register-addin.ps1` 或 `scripts\register-word-addin.ps1`。

## 推荐 draw.io Desktop 版本

桌面模式仍推荐使用 [draw.io Desktop v30.0.4 XML 工具修改版](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1)。

推荐资产：

```text
draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip
```

解压后在插件设置里把 `Desktop Path` / `桌面版路径` 指向：

```text
win-unpacked\draw.io.exe
```

该版本支持通过工具栏直接复制和粘贴 Draw.io XML，便于排查 Office 文件内嵌源数据。

## 验证

- PowerShell 脚本解析检查通过。
- Release 构建通过，0 警告、0 错误。
- 发布包生成通过。
- 发布包内容检查确认包含 PowerPoint/Word 插件和统一注册/注销脚本。
- 发布包临时安装与卸载通过，PowerPoint/Word 注册项均成功写入和清理。
- Word COM Add-in 加载检查通过，`WordAddInConnect=True`。

详细结果见：

- [v1.0.4 E2E 测试报告](./e2e-test-report-v1.0.4.md)
- [v1.0.4 发布证据与日志索引](./release-evidence-v1.0.4.md)
