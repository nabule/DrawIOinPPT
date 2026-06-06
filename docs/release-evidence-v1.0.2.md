# v1.0.2 发布证据与日志索引

[中文](./release-evidence-v1.0.2.md) | [English](./release-evidence-v1.0.2.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本文档汇总 `DrawioPpt v1.0.2` 的程序交付物、真实验证结果和日志位置。

## 1. 交付物

- 版本：`v1.0.2`
- 发布目录：`artifacts\releases\v1.0.2\package`
- 发布压缩包：`artifacts\releases\v1.0.2\DrawioPpt-v1.0.2.zip`
- 安装入口：`install.cmd`
- PowerShell 安装脚本：`scripts\install-release.ps1`
- 卸载入口：`uninstall.cmd`

## 2. 文档

- [v1.0.2 发布说明](./release-notes-v1.0.2.md)
- [v1.0.2 E2E 测试报告](./e2e-test-report-v1.0.2.md)
- [用户手册](./user-guide.md)
- [架构设计](./architecture.md)

## 3. 验证摘要

- Release 构建：通过，0 警告、0 错误
- 本地安装：通过
- PowerPoint COM Add-in 注册：`Greensoft.DrawioPptAddIn`，`LoadBehavior=3`
- PowerPoint 新建/编辑开关测试：通过

## 4. 日志位置

```text
C:\Users\nabul\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

测试日志中记录了本轮新建、编辑、工作文件准备和桌面编辑器启动路径。

## 5. 分发说明

- 对外分发时优先提供 `DrawioPpt-v1.0.2.zip`
- 用户解压后运行 `install.cmd` 即可安装当前用户级 PowerPoint 插件
- 已安装旧版时，重新运行安装脚本即可覆盖到 `v1.0.2`
