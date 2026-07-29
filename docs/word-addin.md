# Word 插件设计与使用说明

[中文](./word-addin.md) | [English](./word-addin.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 1. 目标

Word 版插件的目标是把当前 PowerPoint 插件的 Draw.io 工作流带到 Word Desktop：

- 在当前光标位置插入 Draw.io 图形。
- 从 Draw.io 源生成宽度 1240px、保持源宽高比的 PNG 预览。
- 以 `wdWrapFront` 浮动图片显示；用户选中图片后点击命令才重新进入 Draw.io 编辑。
- 保存后把新的 PNG 预览和 Draw.io XML 回写到原图片。
- 把源 XML 同步保存到 Word 文档的 `Document.CustomXMLParts`，并用图片 `AlternativeText` 保存轻量引用和失败回退。
- 继续支持桌面 draw.io / diagrams.net 和 URL 模式编辑器。

## 2. 当前实现范围

新增项目是 `src/DrawioPpt.WordAddIn/`，它是独立的 Word COM Add-in 宿主。当前为降低一次性改动风险，Word 宿主复用现有 `DrawioPpt.PowerPointAddIn` 程序集中的公共编辑器和 SVG 服务，包括：

- 桌面 draw.io 启动与导出。
- URL 模式 `WebView2 + diagrams.net embed` 编辑器。
- 设置窗口、日志、用户提示。
- SVG 导出与 sanitize、1240px PNG 预览生成、draw.io `content` 元数据补写。

Word 专属代码负责：

- Word COM Add-in 注册和 Ribbon 回调。
- Word 显式双击处理；不监听普通选区变化。
- `InlineShape` / 浮动 `Shape` 图片识别。
- 1240px 等比例 PNG 浮动图片插入和替换。

Word 的展示路径与 PowerPoint 分开：PowerPoint 继续直接插入原始 SVG；Word 正常路径从 Draw.io 源导出宽度 1240px 的 PNG，保持源宽高比，导出明确请求 `--border 0`，并把图片转换为 `wdWrapFront` 浮动 `Shape`。横向和纵向图均在文档可用边界内 `contain`，纵向图回归比例为 `0.25`。重新编辑或刷新采用“先创建并完整配置新图，再删除原图”；失败时清理未完成的新图并保留原图。draw.io Desktop 导出不可用时以 1240px 为基准、按源比例生成轻量 PNG 占位预览，极端纵向图受 1200px 高度上限保护，不回退复杂 SVG；完整 Draw.io XML 仍保存在文档级主存储中，用户选中占位图后点击“重新编辑”仍可恢复编辑。实际 24bpp 不透明 PNG 的透明空白边像素扫描不适用；验收记录独立 SVG 源比例对比和 `--border 0`，不把不支持的像素扫描写成通过。

预览提供器仅在 `%TEMP%\DrawioPpt\word-preview` 下创建临时 Draw.io XML 和 PNG。临时 XML 删除采用 4 次、间隔 50ms 的有限重试；同步清理未完成时安排后台延迟重试，提供器启动时还会清理一小时前遗留的受管目录。
- `Document.CustomXMLParts` 写入、读取和孤儿清理。
- Word 文档路径下 sidecar `.drawio` 重定位。

### 2.1 零选区热路径与显式操作

Word 宿主不订阅 `WindowSelectionChange`，启动时也不读取当前选区、图片属性、`AlternativeText` 或 `Document.CustomXMLParts`。普通选中、拖动、调整大小和重新定位期间，插件不做元数据识别、功能区状态刷新、孤儿清理或自动打开，Word 只执行自身的图片交互。

“重新编辑”“刷新”“绑定”和“清除绑定”始终保持可点击；用户先选中图片，再点击相应命令，插件才读取实时 Word 选区并校验元数据。功能区信息固定提示“选中后点击操作 / 点击按钮时识别”，不再随选择变化。双击图片也属于显式编辑动作，仍可打开受管图形。普通选中、拖动、缩放和重定位不执行插件元数据处理；本版本把正常展示改为 1240px 等比例 PNG 和 `wdWrapFront` 浮动布局，但不能据此声称 Word 原生鼠标拖动已经通过验收。

## 3. 数据策略

Word 没有 PowerPoint `Shape.Tags` 这种直接可用的隐藏标签集合，所以 Word 版采用：

- 图片 `AlternativeText`：正常保存不含 `DrawioXml` 的轻量 envelope，用于快速识别和定位主存储。
- `Document.CustomXMLParts`：保存完整元数据包，保证源 XML 跟随 `.docx`。
- 图片 `Title`：保存可读的图形名称。

轻量 envelope 保留 `formatVersion`、`diagramId`、图形名称、编辑模式、编辑目标、sidecar 路径和更新时间，不保存 `DrawioXml`。读取时先用它拿到 `diagramId`，再从 `Document.CustomXMLParts` 找最新完整包。

如果 `CustomXMLParts.Upsert` 失败，插件会把全量 envelope 留在图片 `AlternativeText` 中，避免源数据丢失。旧版文档如果只在 `AlternativeText` 中保存全量 envelope，首次读取时会自动迁入 `Document.CustomXMLParts`；迁移成功后图片改写为轻量引用，二次读取复用同一主存储部件，不重复迁移。

### 3.1 Word 文件内嵌入方式

Word 版当前同样已经把 Draw.io 源 XML 嵌入 `.docx` 文件内部。主存储是 `Document.CustomXMLParts`，其中保存插件自定义 envelope，包含：

- `diagramId`
- 图形显示名称
- 当前编辑模式
- 编辑目标，可能是桌面程序路径或 URL
- sidecar `.drawio` 路径
- 更新时间
- 经过 `gzip + base64` 压缩的 Draw.io XML

选中图片后点击“重新编辑”时，插件才从图片 `AlternativeText` 解析 `diagramId`，再在 `Document.CustomXMLParts` 中查找最新完整 envelope。保存后，插件会同时更新可见 PNG 预览、图片上的轻量引用和文档级 XML；只有主存储写入失败时才在图片上保留全量回退。

sidecar `.drawio` 在 Word 版里也是编辑缓存和人工备份，不是主存储。把 `.docx` 移到另一台机器后，只要 `Document.CustomXMLParts` 没被清理，插件仍可以从文档内部恢复 Draw.io XML，并在需要桌面版 draw.io 时重新生成 sidecar。

当前没有把 `.drawio` 文件作为 OLE 对象或 Office 附件嵌入 Word。这样可以避免 Office 安全提示、外部文件关联和跨机器打开差异；兼容取舍是：Office 不保证单独复制一张 Word 图片时带走原文档的 `CustomXMLParts`，而正常轻量 `AlternativeText` 又不含 Draw.io XML，因此复制到另一文档后编辑源可能无法恢复。建议复制整份 `.docx`，或事先保存/保留 sidecar `.drawio` 或 Draw.io XML。只有旧版图片或主存储写入失败的回退图片，`AlternativeText` 才可能含全量数据。

### 3.2 推荐桌面编辑器

Word 和 PowerPoint 当前共用同一套插件设置。使用桌面模式时，推荐把 `Desktop Path` 配置为 [draw.io Desktop v30.0.4 XML 工具修改版](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1) 解压目录中的 `win-unpacked\draw.io.exe`。

该版本基于官方 `v30.0.4`，第二行工具栏新增两个彩色 XML 图标按钮：

- 从剪贴板粘贴 draw.io XML 源码并显示为图形。
- 将当前画布复制为 draw.io XML 源码。

这对 Word 插件尤其有用：当你需要核对 `.docx` 内嵌的 Draw.io XML、手工恢复某个图形，或把一段 XML 快速放回桌面编辑器检查时，不需要额外打开源码窗口。下载 `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip` 后完整解压，不要只拷贝单独的 exe。SHA256 为：

```text
B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D
```

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

开发环境注册后打开 Word，会出现 `Draw.io` 功能区：

- `新建`：在当前光标位置插入新的 Draw.io 图形。
- `重新编辑`：编辑当前选中的已绑定图片。
- `刷新`：从绑定的 `.drawio` 工作文件重新生成图片。
- `绑定`：把一个普通图片纳入 Draw.io 管理。
- `清除绑定`：移除绑定元数据，保留图片显示。
- `设置`：复用现有编辑器设置。

Word 不提供“自动打开”入口；这是为了保证选择和拖动过程中没有插件处理。PowerPoint 的自动打开行为不受影响。

## 6. 验证

当前已通过这些真实验证：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-svg-aspect-ratio-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Debug
```

验证覆盖：

- 解决方案 Release x64 构建通过，`0` 个警告、`0` 个错误。
- 真实 Word COM 打开 `.docx`。
- URL 模式创建 Draw.io 图形并回写 Word PNG 预览/XML。
- 保存后重开 Word 文档，再次编辑并回写。
- `Document.CustomXMLParts` 持久化存在。
- 复杂 Draw.io XML 保存在 `CustomXMLParts`，图片 `AlternativeText` 为不含 `DrawioXml` 的轻量引用；保存、关闭、重开后仍成立。
- `CustomXMLParts` 写入失败时全量图片回退可持久化；旧版全量元数据迁移后再次读取保持 XML part ID 稳定。
- Word COM Add-in 能通过 `COMAddIns.Item("Greensoft.DrawioWordAddIn")` 加载，`Connect=True`。
- `word-url-addin-host-e2e.ps1` 在真实 `WINWORD.EXE` 中加载被测 DLL，先选中受管图片，再通过插件的显式命令入口执行“重新编辑”；验证 `ExplicitEditAutomationAvailable=True`、`ExplicitEditCommandInvoked=True`、`configure -> init -> load -> save -> export` 回写、用户级 WebView2 目录以及日志中没有 `E_ACCESSDENIED`。
- `word-preview-image-provider-e2e.ps1` 验证正常 PNG 的签名、1240px 宽度和源比例，以及 2:1 样本在导出失败时生成的 1240×620 等比例轻量 PNG 占位预览；同时验证安全临时路径以及即时源文件和预览文件清理，不回退复杂 SVG。有限重试、后台延迟和启动清理属于实现机制，该 E2E 未通过故障注入验证这些机制。
- `word-svg-aspect-ratio-e2e.ps1` 在真实 Word 中验证 2:1 PNG 新建和替换后均保持 2:1、始终为 `wdWrapFront` 浮动图片；纵向图 `VerticalRatio=0.25`、`VerticalContained=True`；失败替换输出 `FailedReplacementPreservedOriginal=True`，证明原图未被提前删除。
- Word 自动化脚本结束后没有残留 `WINWORD.EXE`。

`word-url-e2e.ps1` 不会强制关闭用户已有 Word 进程；如果检测到 Word 正在运行，会直接中止，避免影响未保存文档。为隔离测试自身创建的宿主，它会在测试期间暂时禁用已注册插件的自动加载，并在 finally 中恢复原 `LoadBehavior`；运行期间不要启动 Word 或并行执行其他 Word 自动化。

`word-url-addin-host-e2e.ps1` 同样要求 Word 已关闭，但它临时注册指定 DLL，运行完成后恢复 Word 加载项、COM 类注册和插件设置。它只会结束由自身 helper 报告且启动时间匹配的 Word 进程，不会按“测试期间新出现的全部 Word 进程”清理。URL 模式的 WebView2 profile 固定为 `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`，不需要 Office 安装目录写权限。
