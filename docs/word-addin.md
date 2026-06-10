# Word 插件设计与使用说明

[中文](./word-addin.md) | [English](./word-addin.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 1. 目标

Word 版插件的目标是把当前 PowerPoint 插件的 Draw.io 工作流带到 Word Desktop：

- 在当前光标位置插入 Draw.io 图形。
- 以 SVG 图片形式显示，尽量保持矢量清晰度。
- 选中已绑定图片后重新进入 Draw.io 编辑。
- 保存后把新的 SVG 和 Draw.io XML 回写到原图片。
- 把源 XML 同步保存到 Word 文档的 `Document.CustomXMLParts`，并用图片 `AlternativeText` 做回退。
- 继续支持桌面 draw.io / diagrams.net 和 URL 模式编辑器。

## 2. 当前实现范围

新增项目是 `src/DrawioPpt.WordAddIn/`，它是独立的 Word COM Add-in 宿主。当前为降低一次性改动风险，Word 宿主复用现有 `DrawioPpt.PowerPointAddIn` 程序集中的公共编辑器和 SVG 服务，包括：

- 桌面 draw.io 启动与导出。
- URL 模式 `WebView2 + diagrams.net embed` 编辑器。
- 设置窗口、日志、用户提示。
- SVG 预览、SVG sanitize、draw.io `content` 元数据补写。

Word 专属代码负责：

- Word COM Add-in 注册和 Ribbon 回调。
- Word 选区监听与双击处理。
- `InlineShape` / 浮动 `Shape` 图片识别。
- SVG 图片插入和替换。
- `Document.CustomXMLParts` 写入、读取和孤儿清理。
- Word 文档路径下 sidecar `.drawio` 重定位。

## 3. 数据策略

Word 没有 PowerPoint `Shape.Tags` 这种直接可用的隐藏标签集合，所以 Word 版采用：

- 图片 `AlternativeText`：保存压缩后的 Draw.io 元数据包，作为快速识别和兼容回退。
- `Document.CustomXMLParts`：保存完整元数据包，保证源 XML 跟随 `.docx`。
- 图片 `Title`：保存可读的图形名称。

读取时优先用图片 `AlternativeText` 拿到 `diagramId`，再从 `Document.CustomXMLParts` 找最新完整包；如果只找到 `AlternativeText`，会自动补写到文档级 CustomXMLParts。

## 4. 发布包安装方式

从发布包安装时，先解压 `DrawioPpt-<version>.zip`，关闭正在运行的 PowerPoint 和 Word，然后在解压目录执行：

```powershell
.\install.cmd
```

安装脚本会把文件复制到当前用户目录，并同时注册 PowerPoint 与 Word 两个 COM Add-in。默认安装位置是：

```text
%LOCALAPPDATA%\Greensoft\DrawioPpt
```

如果只想验证某个临时目录，也可以显式指定安装根目录：

```powershell
.\install.cmd -InstallRoot C:\tmp\DrawioPptInstallCheck
```

卸载时同样先关闭 PowerPoint 和 Word，再执行：

```powershell
.\uninstall.cmd
```

如果升级安装时提示旧目录中 `WebView2Loader.dll` 一类文件被占用，通常是旧的 WebView2 子进程还没有退出。先关闭 Office、插件设置窗口和 URL 编辑窗口；如果仍被占用，重启系统后再次执行安装脚本。安装脚本不会自动结束这些进程，避免影响其它 WebView2 应用。

安装完成后打开 Word，会出现 `Draw.io` 功能区；打开 PowerPoint 时原有 PowerPoint 插件也会继续可用。

## 5. 开发构建与注册

构建：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

注册 Word 插件：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\register-word-addin.ps1
```

卸载 Word 插件：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\unregister-word-addin.ps1
```

开发环境注册后打开 Word，会出现 `Draw.io` 功能区。入口与 PowerPoint 版保持一致：

- `新建`：在当前光标位置插入新的 Draw.io 图形。
- `重新编辑`：编辑当前选中的已绑定图片。
- `刷新`：从绑定的 `.drawio` 工作文件重新生成图片。
- `绑定`：把一个普通图片纳入 Draw.io 管理。
- `清除绑定`：移除绑定元数据，保留图片显示。
- `自动打开`：选中已绑定图片时自动进入编辑。
- `设置`：复用现有编辑器设置。

## 6. 验证

当前已通过这些真实验证：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Debug -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Debug
```

验证覆盖：

- 解决方案 Debug 构建通过。
- 真实 Word COM 打开 `.docx`。
- URL 模式创建 Draw.io 图形并回写 SVG/XML。
- 保存后重开 Word 文档，再次编辑并回写。
- `Document.CustomXMLParts` 持久化存在。
- Word COM Add-in 能通过 `COMAddIns.Item("Greensoft.DrawioWordAddIn")` 加载，`Connect=True`。

`word-url-e2e.ps1` 不会强制关闭用户已有 Word 进程；如果检测到 Word 正在运行，会直接中止，避免影响未保存文档。
