# Full E2E Test Report (v1.0.4)

[中文](./e2e-test-report-v1.0.4.md) | [English](./e2e-test-report-v1.0.4.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本报告记录 `DrawioPpt v1.0.4` 的真实发布验证结果，重点验证发布包生成、统一 Office 注册脚本、PowerPoint/Word 安装链路，以及 Word COM Add-in 加载。

## 1. 测试环境

- 操作系统：Windows
- Office 宿主：Microsoft PowerPoint Desktop、Microsoft Word Desktop
- 构建配置：`Release|x64`
- 注册方式：当前用户级 COM Add-in 注册
- 安装根目录：临时目录 `%TEMP%\DrawioPptReleaseV104-<guid>`

## 2. 执行命令

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.4 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.4\package\scripts\install-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV104-<guid>
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.4\package\scripts\uninstall-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV104-<guid>
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Release
```

另运行了 PowerShell AST 解析检查和包内容检查，覆盖新增的统一注册/注销脚本。

## 3. 构建和打包结果

- Release 构建通过。
- 构建结果：0 警告、0 错误。
- 发布目录已生成：`artifacts\releases\v1.0.4\package`
- 发布压缩包已生成：`artifacts\releases\v1.0.4\DrawioPpt-v1.0.4.zip`
- 包内容检查确认没有 PDB、日志或截图资产。

## 4. 安装验证

- 发布包安装脚本成功调用 `scripts\register-office-addins.ps1`。
- PowerPoint COM Add-in 注册项成功写入：`Greensoft.DrawioPptAddIn`。
- Word COM Add-in 注册项成功写入：`Greensoft.DrawioWordAddIn`。
- PowerPoint/Word 注册项 `LoadBehavior` 均为 `3`。
- 发布包卸载脚本成功调用 `scripts\unregister-office-addins.ps1`。
- 卸载后 PowerPoint/Word 注册项均已清理。

临时安装验证后，已恢复本机默认安装目录中的 Office 插件注册。

## 5. Word 加载验证

`scripts\word-addin-load-check.ps1 -Configuration Release` 验证结果：

```text
WordAddInConnect=True
```

这说明当前 Release 构建的 Word COM Add-in 可被 Word 加载并连接。

## 6. 已知限制

本次发布聚焦安装脚本和打包清单，不修改 Draw.io 编辑协议、SVG 生成逻辑或 Word 图片回写逻辑。
