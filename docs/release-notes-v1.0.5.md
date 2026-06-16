# v1.0.5

[中文](./release-notes-v1.0.5.md) | [English](./release-notes-v1.0.5.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.5` 是面向 Word 安装验收的修正版。`v1.0.4` 的发布包已经包含 Word 注册入口，但旧的 `word-addin-load-check.ps1` 在测试结束时会注销 Word 插件，导致跑完检查后默认环境变成“Word 没装上”。本版本修复该测试脚本，并补上默认安装目录和 Word 实际 COM 加载验收。

## 主要变化

- 新增 `scripts\verify-office-install.ps1`，用于验收发布包安装结果。
- 修复 `scripts\word-addin-load-check.ps1`：测试前记录已有 Word 注册，测试后恢复原注册，不再把默认安装卸掉。
- `install-release.ps1` 在注册 PowerPoint 和 Word 后会立即执行非交互验收，确认两个宿主的 DLL、COM 注册和 `CodeBase` 指向正确。
- 完整 E2E 脚本加入安装后验收步骤，会实际创建 Word/PowerPoint COM 应用并连接 `COMAddIns`。
- 发布包清单加入 `scripts\verify-office-install.ps1`。
- README、安装说明、架构说明和用户手册均更新到 `v1.0.5`，并补充安装后验收命令。

## 根因说明

本次复查发现：默认安装目录一度仍是旧包内容，且当前用户级 PowerPoint/Word 注册项都不存在。直接从 GitHub 下载 `v1.0.4` 资产并运行 `install.cmd` 后，Word 可以注册并加载，说明线上资产本身包含 Word 注册脚本。

进一步复现确认，旧的 `word-addin-load-check.ps1` 会在 `finally` 中调用 `unregister-word-addin.ps1`。因此只要运行该检查脚本，Word 注册就会被删除。`v1.0.5` 改为恢复检查前的注册状态，并通过安装验收脚本把默认安装路径上的实际 Word COM 加载固化为发布前检查项。

## 安装后验收

安装完成后可以运行：

```powershell
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

期望输出包含：

```text
Word registration verified: Greensoft.DrawioWordAddIn
Word.Application COM load verified: Greensoft.DrawioWordAddIn
PowerPoint.Application COM load verified: Greensoft.DrawioPptAddIn
DrawioPpt Office installation verified.
```

## 验证

- `word-addin-load-check.ps1` 保留 Word 注册的回归检查：先失败后通过。
- 新增安装验收覆盖检查：先失败后通过。
- 默认安装目录验收通过，Word 和 PowerPoint 均可通过 Office COM 加载。
- 从 GitHub Release 下载 `v1.0.4` 资产复测，确认线上资产包含统一注册脚本，且默认安装后 Word 可加载。
- `v1.0.5` Release 构建、打包、默认安装、安装验收和 Word 加载检查均作为发布门槛执行。

详细结果见：

- [v1.0.5 E2E 测试报告](./e2e-test-report-v1.0.5.md)
- [v1.0.5 发布证据与日志索引](./release-evidence-v1.0.5.md)
