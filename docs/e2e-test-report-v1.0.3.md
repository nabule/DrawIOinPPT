# Full E2E Test Report (v1.0.3)

[中文](./e2e-test-report-v1.0.3.md) | [English](./e2e-test-report-v1.0.3.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本报告记录 `DrawioPpt v1.0.3` 的真实发布验证结果，重点验证发布包生成、PowerPoint/Word 安装脚本链路，以及文档中新增的 draw.io Desktop XML 工具版配置说明。

## 1. 测试环境

- 操作系统：Windows
- Office 宿主：Microsoft PowerPoint Desktop、Microsoft Word Desktop
- 构建配置：`Release|x64`
- 注册方式：当前用户级 COM Add-in 注册
- 推荐外部编辑器：`draw.io Desktop v30.0.4 XML 工具修改版`

## 2. 执行命令

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.3 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.3\package\scripts\install-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV103-2814525ef98e4c5a9e03c7cdb3e30c3e
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.3\package\scripts\uninstall-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV103-2814525ef98e4c5a9e03c7cdb3e30c3e
```

## 3. 构建和打包结果

- Release 构建通过。
- 构建结果：0 警告、0 错误。
- 发布目录已生成：`artifacts\releases\v1.0.3\package`
- 发布压缩包已生成：`artifacts\releases\v1.0.3\DrawioPpt-v1.0.3.zip`
- 包内容检查确认没有 PDB、日志或截图资产。

## 4. 安装验证

- 发布包安装脚本成功注册 PowerPoint COM Add-in。
- 发布包安装脚本成功注册 Word COM Add-in。
- 发布包卸载脚本成功注销 PowerPoint COM Add-in。
- 发布包卸载脚本成功注销 Word COM Add-in。

临时安装验证后，已恢复本机默认安装目录中的 PowerPoint/Word 插件注册。

## 5. draw.io Desktop XML 工具版说明验证

本版本文档明确推荐：

- 版本：`draw.io Desktop v30.0.4 XML 工具修改版`
- 资产：`draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip`
- 运行路径：`win-unpacked\draw.io.exe`
- SHA256：`B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D`

文档已说明该版本的 XML 工具按钮能力：

- 从剪贴板粘贴 draw.io XML 源码并显示为图形。
- 将当前画布复制为 draw.io XML 源码。

## 6. 已知限制

本次发布主要是 Word 插件交付、文档收口和桌面编辑器配置说明更新；未修改 URL 编辑器保存协议和 SVG 生成逻辑。
