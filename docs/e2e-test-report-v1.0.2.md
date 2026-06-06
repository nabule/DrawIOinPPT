# Full E2E Test Report (v1.0.2)

[中文](./e2e-test-report-v1.0.2.md) | [English](./e2e-test-report-v1.0.2.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本报告记录 `DrawioPpt v1.0.2` 的真实发布验证结果，重点验证新增的图形信息弹窗配置项、安装包部署和 PowerPoint 宿主行为。

## 1. 测试环境

- 操作系统：Windows
- PowerPoint：Microsoft PowerPoint Desktop
- 构建配置：`Release|x64`
- 安装目录：`C:\Users\nabul\AppData\Local\Greensoft\DrawioPpt`
- 注册方式：当前用户级 PowerPoint COM Add-in 注册

## 2. 执行命令

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.2 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.2\package\scripts\install-release.ps1
```

## 3. 构建和安装结果

- Release 构建通过
- 构建结果：0 警告、0 错误
- 发布目录已生成：`artifacts\releases\v1.0.2\package`
- 发布压缩包已生成：`artifacts\releases\v1.0.2\DrawioPpt-v1.0.2.zip`
- 安装脚本已成功注册 `Greensoft.DrawioPptAddIn`
- 注册表 `LoadBehavior=3`

## 4. PowerPoint 宿主验证

通过 PowerPoint COM 自动化启动真实 PowerPoint 宿主，并调用插件主流程验证新配置。

```text
Case   ShowDialogSetting ShapesBefore ShapesAfter ClosedDialogs
Create True              0            1           1
Create False             0            1           0
Edit   True              0            1           1
Edit   False             0            1           0
```

结论：

- 新建 Draw.io 时，配置开启会显示图形信息弹窗，关闭不会显示
- 编辑已有 Draw.io 时，配置开启会显示图形信息弹窗，关闭不会显示
- 四个用例均成功插入或识别插件管理的图形
- 测试后已恢复本机正常配置：桌面编辑器路径为 `C:\Program Files\draw.io\draw.io.exe`

## 5. 已知限制

`scripts\url-editor-smoke.ps1` 在当前自动化环境中曾出现 WinForms/WebView2 窗口未自动退出导致的超时。本次改动不修改 URL 编辑器保存协议，发布验证重点放在新增设置、桌面模式新建和编辑入口。
