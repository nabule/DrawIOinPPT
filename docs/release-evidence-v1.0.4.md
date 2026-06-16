# v1.0.4 发布证据与日志索引

[中文](./release-evidence-v1.0.4.md) | [English](./release-evidence-v1.0.4.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本文档汇总 `DrawioPpt v1.0.4` 的程序交付物、真实验证结果和发布信息。

## 1. 交付物

- 版本：`v1.0.4`
- 发布目录：`artifacts\releases\v1.0.4\package`
- 发布压缩包：`artifacts\releases\v1.0.4\DrawioPpt-v1.0.4.zip`
- 安装入口：`install.cmd`
- 统一注册脚本：`scripts\register-office-addins.ps1`
- 统一注销脚本：`scripts\unregister-office-addins.ps1`
- PowerPoint 插件：`bin\DrawioPpt.PowerPointAddIn.dll`
- Word 插件：`bin\DrawioPpt.WordAddIn.dll`

## 2. 文档

- [v1.0.4 发布说明](./release-notes-v1.0.4.md)
- [v1.0.4 E2E 测试报告](./e2e-test-report-v1.0.4.md)
- [用户手册](./user-guide.md)
- [安装与联调说明](./installation.md)
- [Word 插件设计与使用说明](./word-addin.md)
- [架构设计](./architecture.md)

## 3. 验证摘要

- 新增统一安装入口覆盖测试：先失败后通过。
- PowerShell 脚本解析检查：通过。
- Release 构建：通过，0 警告、0 错误。
- 发布打包：通过。
- 包内容检查：包含统一注册/注销脚本，未包含 PDB、日志或截图资产。
- 临时安装：通过，PowerPoint/Word COM Add-in 均注册成功。
- 临时卸载：通过，PowerPoint/Word COM Add-in 均注销成功。
- Word 加载检查：通过，`WordAddInConnect=True`。

## 4. 推荐外部编辑器

- 名称：`draw.io Desktop v30.0.4 XML 工具修改版`
- Release：https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1
- 资产：`draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip`
- SHA256：`B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D`
- 运行方式：完整解压后运行 `win-unpacked\draw.io.exe`

## 5. 分发说明

- 对外分发优先提供 `DrawioPpt-v1.0.4.zip`。
- 用户解压后运行 `install.cmd`，会同时注册 PowerPoint 和 Word 插件。
- 已安装旧版时，关闭 PowerPoint/Word 后重新运行安装脚本即可覆盖。
- 仓库开发调试时使用 `scripts\register-office-addins.ps1` 注册双宿主插件。
