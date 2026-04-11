# DrawioPpt 用户手册

## 1. 产品说明

DrawioPpt 是一个面向 PowerPoint Desktop 的原生插件，用来把 Draw.io / diagrams.net 图形以可编辑、可回写的方式带进 PPT。

核心目标：

- 在 PPT 中插入矢量图形
- 需要时重新编辑
- 保存后回写到原图形
- 尽量让原始 Draw.io XML 跟着 PPT 一起走

建议先配合以下文档一起使用：

- [安装与联调说明](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/installation.md)
- [完整 E2E 测试报告](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/e2e-test-report-v0.6.0.md)
- [待优化功能点](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/optimization-backlog.md)

## 2. 当前可用功能

### 已可用

- 在当前幻灯片插入新的 Draw.io 图形
- 以 SVG 形式显示图形，保持矢量显示
- 识别插件管理的图形
- 双击已绑定图形直接进入编辑
- 支持桌面版 `draw.io` / `diagrams.net` 编辑
- 支持 URL 模式编辑
- 支持保存后回写 SVG 与 XML
- 支持把原始 XML 存入 `Presentation.CustomXMLParts`
- 支持旧版 `AlternativeText` 元数据自动迁移
- 支持日志记录与 URL 模式自检
- 支持在 URL 模式自动启用 `simpleLabels`，改善 PowerPoint 中 SVG 标签缩放清晰度

### 已提供的辅助能力

- 设置窗口中的 `Detect` 自动探测本地编辑器路径
- 设置窗口中的 `Test URL` 测试 URL 编辑器链路
- 桌面模式自动监听 `.drawio` 文件变化并尝试刷新
- 生成的 SVG 额外补写 draw.io `content` 元数据

## 3. 仍在持续打磨的部分

这些能力已经能用，但还没有做到最终理想状态：

- SVG 替换后对动画、复杂动作和特殊格式的保真仍有提升空间
- URL 模式目前主要依赖 embed 协议，具体行为会受目标 URL 部署方式影响
- sidecar `.drawio` 文件路径在“另存为/移动目录”后的自动重定位还可以更智能
- 多图形批量操作和更丰富的诊断界面还没有做

## 4. 安装方式

推荐先阅读：

- [安装与联调说明](C:/Users/nabul/Desktop/greensoft/code/drawioppt/docs/installation.md)

如果你拿到的是发布包，最简单的安装方式是：

1. 解压发布包
2. 双击 `install.cmd`
3. 打开 PowerPoint
4. 确认插件已经出现在功能区中

如果需要卸载：

1. 关闭 PowerPoint
2. 双击发布包里的 `uninstall.cmd`

## 5. 首次配置

打开插件设置后，重点关注以下项：

- `Editor Mode`
  - `Desktop`：调用本地 `draw.io.exe`
  - `Url`：调用 URL 编辑器
- `Desktop Path`
  - 指向 `draw.io` 或 `diagrams.net Desktop`
- `Editor URL`
  - 指向支持 draw.io embed 协议的编辑器地址
- `Auto open on selection`
  - 选中已绑定图形时自动打开编辑器
- `Auto update on save`
  - 保存 `.drawio` 后自动刷新图形
- `Use MS Office-compatible SVG labels`
  - 在 URL 模式下自动向 draw.io 嵌入编辑器下发 `{ "simpleLabels": true }`
- `Keep sidecar file next to PPT`
  - 在 PPT 同目录保留 `.drawio` 工作文件

推荐配置顺序：

1. 先决定用 `Desktop` 还是 `Url`
2. 如果是 `Desktop`，先点 `Detect`
3. 如果是 `Url`，先点 `Test URL`
4. 再决定是否打开 `Auto open on selection`
5. 最后按需要决定是否保留 sidecar 文件

## 6. 基本用法

### 新建图形

1. 打开一个 PPT
2. 点击插件中的“新建 Draw.io 图形”
3. 进入桌面编辑器或 URL 编辑器
4. 保存后，图形会回到当前幻灯片

### 编辑已有图形

1. 选中一个插件管理的图形
2. 双击图形，或点击“编辑”
3. 修改并保存
4. 插件会回写当前图形

如果启用了 `Auto open on selection`，选中受管理图形时也可能自动打开编辑器。

### 刷新图形

适合桌面模式下修改了 `.drawio` 文件但未自动回写的情况：

1. 选中图形
2. 点击“刷新”

### 绑定普通图形

适合把一个现有图形纳入插件管理：

1. 选中图形
2. 点击“绑定”
3. 插件会为该图形写入 Draw.io 元数据

### 清除绑定

1. 选中图形
2. 点击“清除绑定”
3. 图形会保留显示，但不再由插件管理

## 7. 两种编辑模式的区别

### Desktop 模式

优点：

- 本地编辑体验更稳定
- 更适合企业内网
- 与本地 `.drawio` 工作文件联动更自然

注意：

- 需要安装桌面版 Draw.io
- 首次配置要确保 `Desktop Path` 正确
- 如果想让导出的 SVG 标签在 PowerPoint 中缩放更清晰，建议在 draw.io Desktop 中手动设置 `simpleLabels`

### URL 模式

优点：

- 不依赖本地桌面编辑器
- 易于接入私有化 URL
- 可先用 `Test URL` 自检
- `v0.6.1` 已自动注入 Office 兼容 SVG 标签配置

注意：

- 目标 URL 必须支持 draw.io embed 协议
- 会受到网络、代理、私有部署差异影响

URL 模式下的 SVG 标签清晰度优化：

1. 插件会自动给 embed URL 加上 `configure=1`
2. 在编辑器握手阶段自动发送 `{ "simpleLabels": true }`
3. 导出的 SVG 在 PowerPoint 中缩放时，文字边缘通常会更干净

## 8. 数据保存策略

DrawioPpt 目前使用三层策略保存源数据：

1. `Presentation.CustomXMLParts`
2. `Shape.Tags`
3. `Shape.AlternativeText` 兼容回退

同时，生成出的 SVG 还会补写 draw.io `content` 元数据。

这意味着：

- PPT 文件本身会保留主要源数据
- 单个 SVG 文件也带一份可恢复信息
- 旧图形在再次读取时可自动迁移到新存储方案

## 9. 日志与排障

默认日志位置：

```text
C:\Users\<你的用户名>\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

常见排障建议：

- URL 模式先点 `Test URL`
- 桌面模式先确认 `Desktop Path`
- 自动刷新不工作时检查 `.drawio` 文件是否真的保存
- 发生错误时优先看日志里的最近事件

## 10. 已知限制

- 当前只支持 PowerPoint Desktop，不支持 WPS
- 暂未提供 MSI/Setup Wizard 形式的安装器，当前以脚本安装为主
- 动画和复杂格式保真仍不是最终形态
- 多人同时编辑同一 PPT 的冲突处理还没有专门优化
- `simpleLabels` 目前只对 URL 模式由插件自动控制，桌面模式仍需用户在 draw.io Desktop 中手工设置

## 11. 当前建议的使用边界

更适合当前版本的场景：

- 单人编辑的 PowerPoint 文档
- 需要把源 XML 跟 PPT 一起保存的内部资料
- 桌面 PowerPoint 环境
- 已能接受脚本式安装的团队

仍建议谨慎使用的场景：

- 动画和复杂动作非常重的演示文稿
- 多人频繁同时修改同一 PPT
- 需要统一企业级安装器的受管环境
