# Full E2E Test Report (v1.0.5)

[中文](./e2e-test-report-v1.0.5.md) | [English](./e2e-test-report-v1.0.5.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本报告记录 `DrawioPpt v1.0.5` 的真实验证结果，重点验证默认安装路径、Word 注册、Word 实际 COM 加载，以及新增安装验收脚本。

## 1. 测试环境

- 操作系统：Windows
- Office 宿主：Microsoft PowerPoint Desktop、Microsoft Word Desktop
- 构建配置：`Release|x64`
- 注册方式：当前用户级 COM Add-in 注册
- 默认安装目录：`%LOCALAPPDATA%\Greensoft\DrawioPpt`

## 2. 已复现的问题

复查开始时，默认安装目录仍是旧包内容，且以下注册项不存在：

- `HKCU\Software\Microsoft\Office\PowerPoint\Addins\Greensoft.DrawioPptAddIn`
- `HKCU\Software\Microsoft\Office\Word\Addins\Greensoft.DrawioWordAddIn`
- `HKCU\Software\Classes\Greensoft.DrawioWordAddIn`
- `HKCU\Software\Classes\CLSID\{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}`

这与用户观察到的“Word 没装上”一致。随后复现确认，旧的 `scripts\word-addin-load-check.ps1` 会在测试结束时注销 Word 注册，是导致测试后 Word 消失的直接原因。

## 3. 执行命令

```powershell
.\artifacts\releases\v1.0.4\package\install.cmd
powershell -ExecutionPolicy Bypass -File .\scripts\verify-office-install.ps1
gh release download v1.0.4 --repo nabule/DrawIOinPPT --pattern "DrawioPpt-v1.0.4.zip"
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.5 -Configuration Release -Platform x64 -SkipBuild
```

最终发布前还会执行：

```powershell
.\artifacts\releases\v1.0.5\package\install.cmd
powershell -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Release
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

## 4. 关键验证结果

- GitHub 上的 `v1.0.4` 发布资产 SHA256 与本地一致：`9A253FAD59689D76F43002BA1CCBA7C841A78EAAD4BA8164ABE0BE9EB4066197`。
- 从 GitHub 下载并解压的 `v1.0.4` 包包含 `register-office-addins.ps1`。
- 从 GitHub 下载包运行 `install.cmd` 后，Word 注册项写入成功。
- Word `COMAddIns.Item("Greensoft.DrawioWordAddIn")` 可找到并连接，`Connect=True`。
- 旧 `word-addin-load-check.ps1` 会删除 Word 注册，回归测试先失败；修复后同一测试通过，`CodeBase` 保持指向默认安装目录。
- Word `Resiliency\DisabledItems` 中存在的禁用项不是 DrawioWord，而是 Microsoft Word 预览器。
- 新增 `verify-office-install.ps1` 后，默认安装目录验收通过。

验收脚本输出：

```text
PowerPoint registration verified: Greensoft.DrawioPptAddIn
Word registration verified: Greensoft.DrawioWordAddIn
Word.Application COM load verified: Greensoft.DrawioWordAddIn
PowerPoint.Application COM load verified: Greensoft.DrawioPptAddIn
DrawioPpt Office installation verified.
```

## 5. 结论

`v1.0.5` 修复会注销 Word 的加载检查脚本，并把默认安装路径和 Word 实际 COM 加载纳入发布门槛。后续如果 Word 未注册、`CodeBase` 指错、Office 禁用了插件，或 `COMAddIns` 无法连接，验收脚本会直接失败。
