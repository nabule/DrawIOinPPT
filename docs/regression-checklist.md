# 回归清单

## 1. 构建与注册

- `scripts/dev-check.ps1` 通过
- `scripts/build.ps1` 通过
- `scripts/register-addin.ps1` 可完成注册
- PowerPoint 中 `Greensoft.DrawioPptAddIn` 可正常加载

## 2. Ribbon 与选择状态

- Ribbon 可正常显示
- 未选中图形时状态为“未选中图形”
- 选中普通图形时状态为“当前图形未绑定”
- 选中已绑定图形时状态为“已识别”

## 3. 桌面模式

- 可自动探测本地 `draw.io` 路径
- 可新建图形
- 可打开桌面编辑器
- 未保存的空白 PPT 中新建图形时，sidecar 会落到绝对临时路径，不会生成相对 `drawio\*.drawio`
- 保存 `.drawio` 后 PPT 中 SVG 自动刷新
- 手动刷新按钮可用

## 4. URL 模式

- 设置窗口 `Test URL` 可通过
- 可打开已有图形
- `init -> load -> save -> export` 可走通
- SVG 与 XML 可回写到 PPT
- 新建图形后保存 PPT、关闭并重新打开，仍可再次编辑并回写
- 日志中可看到 URL 模式关键事件

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
