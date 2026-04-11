# 安装与联调说明

## 1. 环境前提

- Windows
- Microsoft PowerPoint Desktop x64
- .NET Framework 4.8 Developer Pack
- WebView2 Runtime
- 可选：`draw.io` / `diagrams.net Desktop`

建议先运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
```

## 2. 面向最终用户的安装

如果你拿到的是发布包 `DrawioPpt-v1.0.0.zip`，推荐按下面流程安装：

1. 解压 zip 到一个本地目录
2. 确认 PowerPoint 当前没有运行
3. 双击 `install.cmd`
4. 等待脚本提示安装成功
5. 打开 PowerPoint，确认功能区里出现 DrawioPpt

发布包内的核心安装文件包括：

- `install.cmd`
- `uninstall.cmd`
- `scripts\install-release.ps1`
- `scripts\uninstall-release.ps1`
- `scripts\register-addin.ps1`
- `scripts\unregister-addin.ps1`
- `docs\release-notes-v1.0.0.md`
- `docs\e2e-test-report-v1.0.0.md`
- `docs\release-evidence-v1.0.0.md`
- `logs\`

默认安装位置：

```text
C:\Users\<你的用户名>\AppData\Local\Greensoft\DrawioPpt
```

卸载方式：

1. 关闭 PowerPoint
2. 双击 `uninstall.cmd`
3. 确认 PowerPoint 不再自动加载该插件

## 3. 面向开发的仓库安装

如果你是在仓库里联调开发，请先构建，再注册当前构建输出。

## 4. 构建

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

默认输出：

- `src\DrawioPpt.Core\bin\x64\Debug\DrawioPpt.Core.dll`
- `src\DrawioPpt.PowerPointAddIn\bin\x64\Debug\DrawioPpt.PowerPointAddIn.dll`

## 5. 注册插件

当前脚本使用当前用户级 COM 注册，无需管理员权限。

注册：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-addin.ps1
```

说明：

- 在仓库开发目录中，脚本会自动定位 `src\DrawioPpt.PowerPointAddIn\bin\...`
- 在发布包目录中，脚本会自动定位同级 `bin\DrawioPpt.PowerPointAddIn.dll`

卸载：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister-addin.ps1
```

## 6. 首次启动建议

1. 打开 PowerPoint。
2. 确认 `Greensoft.DrawioPptAddIn` 已加载。
3. 打开插件设置。
4. 根据实际需要配置：
   - `Desktop Path`
   - `Editor URL`
   - `Auto update on save`
   - `Keep sidecar file next to PPT`

## 7. 安装后建议的验收顺序

推荐按下面顺序确认安装成功：

1. 插件能正常加载
2. 设置窗口能打开
3. `Desktop` 模式可探测本地 `draw.io.exe`
4. `Url` 模式的 `Test URL` 能通过
5. 能创建一个新图形并成功回写

如果你需要完整自动化验收，可以运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1
```

完整测试报告：

- [v1.0.0 E2E 测试报告](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/e2e-test-report-v1.0.0.md)

## 8. 桌面模式验证

1. 在设置中把 `Editor Mode` 设为 `Desktop`。
2. 检查 `Desktop Path` 是否指向 `draw.io.exe` 或 `diagrams.net.exe`。
3. 新建图形。
4. 在桌面编辑器保存后，确认 PPT 中 SVG 自动刷新。

补充说明：

- `v1.0.0` 已在 `Url` 模式下自动启用 MS Office 兼容的 `simpleLabels`。
- `Desktop` 模式调用的是外部桌面版 draw.io，插件目前不能替你改写该桌面应用的编辑器配置。
- 如果你希望桌面模式导出的 SVG 标签在 PowerPoint 中缩放更清晰，建议在 draw.io Desktop 里手动设置：
  - `Extras > Configuration`
  - 输入 `{ "simpleLabels": true }`
  - 点击 `Apply`

## 9. URL 模式验证

1. 在设置中把 `Editor Mode` 设为 `Url`。
2. 填写 `Editor URL`。
3. 先点击设置窗口中的 `Test URL`。
4. 测试通过后，再在 PPT 中新建或编辑图形。

如果需要跑仓库内的 URL 烟雾测试：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1
```

## 10. 日志位置

默认日志文件：

```text
C:\Users\<你的用户名>\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

适合排查：

- URL 初始化超时
- URL 导出超时
- 宿主回写失败
- 存储迁移与孤儿清理异常

## 11. 常见问题

### URL 窗口打开但没有保存回调

- 先点设置中的 `Test URL`
- 检查日志里的 `init/save/export` 事件
- 确认目标地址支持 draw.io embed 协议

### 桌面模式能打开编辑器但不刷新

- 确认 `.drawio` 文件确实被保存
- 确认 `Auto update on save` 已开启
- 检查日志中是否出现刷新异常

### 移动 PPT 后图形还能不能编辑

- 可以，源 XML 保存在 `Presentation.CustomXMLParts`
- `v1.0.0` 会在再次编辑时优先按当前 `PPT` 路径重建 sidecar；如果你把 sidecar 一起带走，桌面模式刷新会更顺畅

### 为什么 URL 模式里的 SVG 标签缩放更清晰

- `v1.0.0` 会在 URL 嵌入编辑器里自动下发 `{ "simpleLabels": true }`
- 这是 draw.io 官方针对 MS Office SVG 缩放兼容性给出的优化建议

### 安装脚本提示 PowerPoint 正在运行

- 关闭所有 PowerPoint 窗口后重试
- 如果后台还残留 `POWERPNT.EXE`，先结束该进程再安装或卸载

### 安装后没有看到插件

- 先重新打开 PowerPoint
- 再检查当前用户注册是否已写入
- 可重新执行一次 `scripts\register-addin.ps1`
