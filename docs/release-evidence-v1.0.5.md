# v1.0.5 发布证据与日志索引

[中文](./release-evidence-v1.0.5.md) | [English](./release-evidence-v1.0.5.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本文档汇总 `DrawioPpt v1.0.5` 的程序交付物、真实验证结果和发布信息。

## 1. 交付物

- 版本：`v1.0.5`
- 发布目录：`artifacts\releases\v1.0.5\package`
- 发布压缩包：`artifacts\releases\v1.0.5\DrawioPpt-v1.0.5.zip`
- 安装入口：`install.cmd`
- 统一注册脚本：`scripts\register-office-addins.ps1`
- 统一注销脚本：`scripts\unregister-office-addins.ps1`
- 安装验收脚本：`scripts\verify-office-install.ps1`
- PowerPoint 插件：`bin\DrawioPpt.PowerPointAddIn.dll`
- Word 插件：`bin\DrawioPpt.WordAddIn.dll`

## 2. 文档

- [v1.0.5 发布说明](./release-notes-v1.0.5.md)
- [v1.0.5 E2E 测试报告](./e2e-test-report-v1.0.5.md)
- [用户手册](./user-guide.md)
- [安装与联调说明](./installation.md)
- [Word 插件设计与使用说明](./word-addin.md)
- [架构设计](./architecture.md)

## 3. 验证摘要

- 复现默认安装目录未注册 Word 的状态：完成。
- 从 GitHub 下载 `v1.0.4` 资产复测：完成，资产包含统一注册脚本且可默认安装 Word。
- 复现 `word-addin-load-check.ps1` 会删除 Word 注册：完成。
- 修复 `word-addin-load-check.ps1` 后重复测试：通过，Word 注册和 `CodeBase` 保留。
- 新增安装验收覆盖测试：先失败后通过。
- 默认安装目录验收：通过，PowerPoint/Word 注册与 COM 加载均通过。
- Word 禁用项检查：未发现 DrawioWord 被 Office 禁用。
- `v1.0.5` Release 构建、打包、默认安装和验收：作为发布门槛执行。

## 4. 分发说明

- 对外分发优先提供 `DrawioPpt-v1.0.5.zip`。
- 用户解压后运行 `install.cmd`，会同时注册 PowerPoint 和 Word 插件，并执行非交互安装验收。
- 安装后如果用户仍看不到 Word 插件，先运行 `scripts\verify-office-install.ps1`，该脚本会区分未注册、`CodeBase` 指错、Office 禁用项和 COM 加载失败。
