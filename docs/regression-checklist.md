# 回归清单

[中文](./regression-checklist.md) | [English](./regression-checklist.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

## 1. 构建与注册

- `scripts/dev-check.ps1` 通过
- `scripts/build.ps1` 通过
- `scripts/register-office-addins.ps1` 可同时完成 PowerPoint 和 Word 注册
- PowerPoint 中 `Greensoft.DrawioPptAddIn` 可正常加载

## 2. Ribbon 与选择状态

- Ribbon 可正常显示
- Ribbon 分组为“创建 / 当前图形 / 工作流 / 信息”
- 未选中图形时状态为“未选中”
- 选中普通图形时状态为“普通图形，可直接绑定”
- 选中已绑定图形时状态为“已识别为 Draw.io 图形”
- 已绑定图形的主按钮文案会显示为“重新编辑”
- `自动打开` 开关可正确读写 `AutoOpenOnSelection`
- 按钮图标可正常显示，不出现空白占位

## 3. 桌面模式

- 可自动探测本地 `draw.io` 路径
- 可新建图形
- 可打开桌面编辑器
- 未保存的空白 PPT 中新建图形时，sidecar 会落到绝对临时路径，不会生成相对 `drawio\*.drawio`
- 已保存或另存为后的 PPT，再次编辑时 sidecar 会自动切换到当前 PPT 目录下
- 将 PPT 拷贝到另一台机器后，再次编辑时 sidecar 会按当前机器上的 PPT 路径重建
- 保存 `.drawio` 后 PPT 中 SVG 自动刷新
- 手动刷新按钮可用

## 4. URL 模式

- 设置窗口 `Test URL` 可通过
- 可打开已有图形
- `init -> load -> save -> export` 可走通
- SVG 与 XML 可回写到 PPT
- 新建图形后保存 PPT、关闭并重新打开，仍可再次编辑并回写
- 日志中可看到 URL 模式关键事件
- `scripts\word-url-addin-host-e2e.ps1` 在真实 `WINWORD.EXE` 中通过：加载项已连接、选择受管图片后完成回写、`%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2` 存在，新增日志没有 `E_ACCESSDENIED`。

## 5. 存储与迁移

- 新图形会写入 `Shape.Tags`
- 新图形会写入 `Presentation.CustomXMLParts`
- 重新打开 PPT 后可从 `CustomXMLParts` 恢复源 XML 并继续编辑
- 旧版只含 `AlternativeText` 的图形会自动迁移
- 删除图形绑定后，不再保留无引用的孤儿 `CustomXMLPart`
- 生成出的 SVG 包含 draw.io `content` 元数据

## 6. SVG 替换保真

- 替换后位置不跳动
- 替换后尺寸不跳动
- 替换后旋转保留
- 替换后名称保留
- 替换后超链接保留

## 7. 异常与日志

- 桌面路径缺失时有明确提示
- URL 初始化超时有明确提示
- URL 导出超时有明确提示
- 宿主入口异常会落日志并弹出错误框
- 日志文件可持续追加写入

## 8. 注册与卸载

- 卸载脚本可正常移除注册项
- 卸载后 PowerPoint 不再加载该插件

## 9. Word 选区与图片操作

- `scripts\word-selection-sync-test.ps1` 通过，`AddInHost` 不得重新引入选区 `Timer` 或 Tick 回调。
- `scripts\word-selection-event-e2e.ps1` 在真实 Word COM 中通过，验证浮动 `Shape` 的受管图片即时识别、普通图片的现场绑定前提，以及发布 DLL 的无轮询约束。
- 执行 Word E2E 前关闭 Word，且不要与其他 Word 自动化测试并行运行；脚本会还原插件设置并仅清理本次测试创建的自动化 Word 进程。
- 移动、缩放浮动图片后，功能区状态由 Word 事件更新；如 Word 未立即更新显示，直接选中图片后点击“重新编辑”“刷新”“绑定”或“清除绑定”，命令会现场读取当前选区。
