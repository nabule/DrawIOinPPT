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

## 2. 构建

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

默认输出：

- `src\DrawioPpt.Core\bin\x64\Debug\DrawioPpt.Core.dll`
- `src\DrawioPpt.PowerPointAddIn\bin\x64\Debug\DrawioPpt.PowerPointAddIn.dll`

## 3. 注册插件

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

## 4. 首次启动建议

1. 打开 PowerPoint。
2. 确认 `Greensoft.DrawioPptAddIn` 已加载。
3. 打开插件设置。
4. 根据实际需要配置：
   - `Desktop Path`
   - `Editor URL`
   - `Auto update on save`
   - `Keep sidecar file next to PPT`

## 5. 桌面模式验证

1. 在设置中把 `Editor Mode` 设为 `Desktop`。
2. 检查 `Desktop Path` 是否指向 `draw.io.exe` 或 `diagrams.net.exe`。
3. 新建图形。
4. 在桌面编辑器保存后，确认 PPT 中 SVG 自动刷新。

## 6. URL 模式验证

1. 在设置中把 `Editor Mode` 设为 `Url`。
2. 填写 `Editor URL`。
3. 先点击设置窗口中的 `Test URL`。
4. 测试通过后，再在 PPT 中新建或编辑图形。

如果需要跑仓库内的 URL 烟雾测试：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1
```

## 7. 日志位置

默认日志文件：

```text
C:\Users\<你的用户名>\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

适合排查：

- URL 初始化超时
- URL 导出超时
- 宿主回写失败
- 存储迁移与孤儿清理异常

## 8. 常见问题

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
- 如果启用了 sidecar，还需要保证 `.drawio` 工作文件路径可用
