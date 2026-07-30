# DrawioPpt v1.0.9 发布说明

[中文首页](../README.md) | [English](./release-notes-v1.0.9.en.md)

## 这版先解决什么问题

此前安装包已经写入新的 BuildId，但 Word 和 PowerPoint 功能区无法显示它。用户无法判断 Office 是否实际加载了刚升级的插件，只能依赖安装脚本输出或手工检查文件。

## 这版如何解决

- Word 和 PowerPoint 都补齐了 Office Ribbon 所调用的版本回调，因此功能区信息区会显示 `版本：v1.0.9+<git-short-hash>`。
- 设置窗口底部继续显示同一个 BuildId。
- 安装验证现在会启动实际 Word 和 PowerPoint，直接读取两个加载项返回的版本文本，并与安装目录 `BuildInfo.txt` 比对。

## 使用方式

安装 `DrawioPpt-v1.0.9.zip` 后，关闭并重新打开 Word 或 PowerPoint，在 `Draw.io` 功能区的“信息”区查看版本。显示值应与安装程序输出的 BuildId 一致。

## 验证

- 版本显示回归测试覆盖 Word、PowerPoint 的 Ribbon XML、COM 入口回调、版本读取和安装验证链路。
- 完整 Office E2E `13/13 PASS`，标准本地安装的 Word、PowerPoint 都返回与 `BuildInfo.txt` 一致的 BuildId。
