# 安装与联调说明

[中文](./installation.md) | [English](./installation.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 1. 环境前提

- Windows
- Microsoft PowerPoint Desktop x64
- Microsoft Word Desktop x64（使用 Word 插件时需要）
- .NET Framework 4.8 Developer Pack
- WebView2 Runtime
- 可选但推荐：[draw.io Desktop v30.0.4 XML 工具修改版](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1)

建议先运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
```

## 2. 面向最终用户的安装

如果你拿到的是发布包 `DrawioPpt-v1.0.8.zip`，推荐按下面流程安装：

1. 解压 zip 到一个本地目录
2. 确认 PowerPoint 和 Word 当前没有运行
3. 双击 `install.cmd`
4. 等待脚本提示安装成功
5. 打开 PowerPoint 或 Word，确认功能区里出现 `Draw.io`

安装成功输出会包含 `BuildId`，格式为 `v1.0.8+<git-short-hash>`。同一值也会写入发布包根目录和安装目录中的 `BuildInfo.txt` / `PACKAGE.txt`，便于核对当前安装到底来自哪个提交。

发布包内的核心安装文件包括：

- `install.cmd`
- `uninstall.cmd`
- `scripts\install-release.ps1`
- `scripts\uninstall-release.ps1`
- `scripts\register-office-addins.ps1`
- `scripts\unregister-office-addins.ps1`
- `scripts\verify-office-install.ps1`
- `BuildInfo.txt`
- `PACKAGE.txt`
- `scripts\register-addin.ps1`
- `scripts\unregister-addin.ps1`
- `scripts\register-word-addin.ps1`
- `scripts\unregister-word-addin.ps1`
- `docs\release-notes-v1.0.8.md`
- `docs\e2e-test-report-v1.0.8.md`
- `docs\release-evidence-v1.0.8.md`

默认安装位置：

```text
C:\Users\<你的用户名>\AppData\Local\Greensoft\DrawioPpt
```

卸载方式：

1. 关闭 PowerPoint
2. 关闭 Word
3. 双击 `uninstall.cmd`
4. 确认 PowerPoint 和 Word 不再自动加载该插件

## 3. 面向开发的仓库安装

如果你是在仓库里联调开发，请先构建，再注册当前构建输出。

## 4. 构建

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

默认输出：

- `src\DrawioPpt.Core\bin\x64\Debug\DrawioPpt.Core.dll`
- `src\DrawioPpt.PowerPointAddIn\bin\x64\Debug\DrawioPpt.PowerPointAddIn.dll`
- `src\DrawioPpt.WordAddIn\bin\x64\Debug\DrawioPpt.WordAddIn.dll`

构建脚本会同时在三个输出目录写入 `BuildInfo.txt`。插件运行时优先从该文件读取 `BuildId`，因此仓库开发注册和发布包安装都能在设置窗口与功能区信息区看到相同的版本号和 Git short hash。

## 5. 注册插件

当前脚本使用当前用户级 COM 注册，无需管理员权限。

同时注册 PowerPoint 和 Word：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-office-addins.ps1
```

说明：

- 在仓库开发目录中，脚本会自动定位 `src\DrawioPpt.PowerPointAddIn\bin\...` 和 `src\DrawioPpt.WordAddIn\bin\...`
- 在发布包目录中，脚本会自动定位同级 `bin\DrawioPpt.PowerPointAddIn.dll` 和 `bin\DrawioPpt.WordAddIn.dll`
- `install.cmd` / `scripts\install-release.ps1` 内部也会调用统一注册入口，因此最终用户安装会同时覆盖 PowerPoint 和 Word

同时注销 PowerPoint 和 Word：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister-office-addins.ps1
```

如果只需要排查单个宿主，可以使用单宿主脚本：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-addin.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\register-word-addin.ps1
```

安装后验收：

```powershell
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

该脚本会检查 PowerPoint/Word 插件 DLL、当前用户级 COM 注册、`CodeBase` 指向、Office 禁用项，以及 Word/PowerPoint 的 `COMAddIns` 实际加载状态。

## 6. 首次启动建议

1. 打开 PowerPoint。
2. 确认 `Greensoft.DrawioPptAddIn` 已加载；如果使用 Word，也确认 `Greensoft.DrawioWordAddIn` 已加载。
3. 打开插件设置。
4. 根据实际需要配置：
   - `Desktop Path`
   - `Editor URL`
   - `Auto update on save`
   - `Keep sidecar file next to PPT`

## 7. 推荐 draw.io Desktop XML 工具版

桌面模式推荐使用 [draw.io Desktop v30.0.4 XML 工具修改版](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1)。它基于官方 `v30.0.4`，在第二行工具栏新增两个彩色 XML 图标按钮：

- 从剪贴板粘贴 draw.io XML 源码并显示为图形。
- 将当前画布复制为 draw.io XML 源码。

安装和配置步骤：

1. 从 release 页面下载 `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip`。
2. 校验 SHA256：

   ```text
   B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D
   ```

3. 解压 zip 到固定目录，例如 `D:\Tools\draw.io-30.0.4-xml-tools.1\`。
4. 运行 `win-unpacked\draw.io.exe`。这是未签名的 Windows x64 免安装目录版，不要只拷贝单独的 exe。
5. 在 DrawioPpt 插件设置中，把 `Editor Mode` 设为 `Desktop`。
6. 把 `Desktop Path` 指向解压目录下的 `win-unpacked\draw.io.exe`，例如：

   ```text
   D:\Tools\draw.io-30.0.4-xml-tools.1\win-unpacked\draw.io.exe
   ```

如果你需要在插件和 draw.io Desktop 之间手工排查源数据，这个版本的 XML 复制/粘贴按钮会比普通官方桌面版更方便。

## 8. 安装后建议的验收顺序

推荐按下面顺序确认安装成功：

1. 插件能正常加载
2. 设置窗口能打开
3. 设置窗口底部和功能区信息区能看到 `版本：v1.0.8+<git-short-hash>`
4. `Desktop` 模式可探测本地 `draw.io.exe`
5. `Url` 模式的 `Test URL` 能通过
6. 能创建一个新图形并成功回写

如果你需要完整自动化验收，可以运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1
```

完整测试报告：

- [v1.0.8 E2E 测试报告](./e2e-test-report-v1.0.8.md)

## 9. 桌面模式验证

1. 在设置中把 `Editor Mode` 设为 `Desktop`。
2. 检查 `Desktop Path` 是否指向 `draw.io.exe` 或 `diagrams.net.exe`。
3. 新建图形。
4. 在桌面编辑器保存后，确认 PPT 中 SVG 自动刷新。

补充说明：

- `v1.0.5` 已在 `Url` 模式下自动启用 MS Office 兼容的 `simpleLabels`。
- `Desktop` 模式调用的是外部桌面版 draw.io，插件目前不能替你改写该桌面应用的编辑器配置。
- 推荐使用 `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip` 解压后的 `win-unpacked\draw.io.exe`，便于手工复制/粘贴 XML。
- 如果你希望桌面模式导出的 SVG 标签在 PowerPoint 中缩放更清晰，建议在 draw.io Desktop 里手动设置：
  - `Extras > Configuration`
  - 输入 `{ "simpleLabels": true }`
  - 点击 `Apply`

## 10. URL 模式验证

1. 在设置中把 `Editor Mode` 设为 `Url`。
2. 填写 `Editor URL`。
3. 先点击设置窗口中的 `Test URL`。
4. 测试通过后，再在 PPT 中新建或编辑图形。

如果需要跑仓库内的 URL 烟雾测试：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1
```

从 `v1.0.7` 起，URL 编辑器的 WebView2 profile 固定保存在当前用户目录：

```text
%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2
```

因此 Word 和 PowerPoint 不需要、也不应尝试获取 Office 安装目录的写入权限。若曾出现“无法初始化 URL 编辑器，拒绝访问”，安装 `v1.0.7` 后关闭并重新打开 Office，再使用 `Test URL` 验证即可。仓库中可用下列真实 Word 宿主回归验证发布 DLL；执行前必须关闭 Word，脚本会临时替换并最终还原当前用户的 Word 加载项注册和设置：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release
```

## 11. 日志位置

默认日志文件：

```text
C:\Users\<你的用户名>\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

适合排查：

- URL 初始化超时
- URL 导出超时
- 宿主回写失败
- 存储迁移与孤儿清理异常

## 12. 常见问题

### URL 窗口打开但没有保存回调

- 先点设置中的 `Test URL`
- 检查日志里的 `init/save/export` 事件
- 确认目标地址支持 draw.io embed 协议

### URL 编辑器提示“无法初始化”或“拒绝访问”

- 确认安装的是 `v1.0.7` 或更高版本，并完全退出后重新打开 Word/PowerPoint。
- 检查 `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2` 是否可由当前用户创建和写入；不要修改 `C:\Program Files\Microsoft Office` 的权限。
- 若目录可写但仍失败，查看日志中的 `UrlEditor` 记录，并确认 WebView2 Runtime 已安装。

### 桌面模式能打开编辑器但不刷新

- 确认 `.drawio` 文件确实被保存
- 确认 `Auto update on save` 已开启
- 检查日志中是否出现刷新异常

### 移动 PPT 后图形还能不能编辑

- 可以，源 XML 保存在 `Presentation.CustomXMLParts`
- `v1.0.5` 会在再次编辑时优先按当前 `PPT` 路径重建 sidecar；如果你把 sidecar 一起带走，桌面模式刷新会更顺畅

### 为什么 URL 模式里的 SVG 标签缩放更清晰

- `v1.0.5` 会在 URL 嵌入编辑器里自动下发 `{ "simpleLabels": true }`
- 这是 draw.io 官方针对 MS Office SVG 缩放兼容性给出的优化建议

### 安装脚本提示 PowerPoint 正在运行

- 关闭所有 PowerPoint 窗口后重试
- 如果后台还残留 `POWERPNT.EXE`，先结束该进程再安装或卸载

### 安装后没有看到插件

- 先重新打开 PowerPoint
- 再检查当前用户注册是否已写入
- 可重新执行一次 `scripts\register-office-addins.ps1`
