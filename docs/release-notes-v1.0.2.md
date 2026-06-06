# v1.0.2

[中文](./release-notes-v1.0.2.md) | [English](./release-notes-v1.0.2.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

`v1.0.2` 是一个面向编辑体验和发布交付的小版本，重点新增“新建或编辑 Draw.io 时是否显示图形信息弹窗”的配置项，并重新打包为带安装脚本的发布包。

## 主要变化

- 设置窗口新增 `新建或编辑时显示图形信息弹窗`
- 新增内部设置项 `ShowDiagramInfoDialog`，旧配置文件缺少该节点时默认开启
- 新建 Draw.io 图形和重新编辑已有图形时，成功类图形信息弹窗统一受该设置控制
- 错误提示不受该设置影响，仍会正常显示
- README、架构文档和用户手册已同步更新

## 安装与升级

- 发布包包含 `install.cmd`、`uninstall.cmd` 和 `scripts\install-release.ps1`
- 已注册过旧版插件的机器，重新运行发布包中的安装脚本即可切换到 `v1.0.2`
- 默认配置会继续显示图形信息弹窗；如需关闭，可在插件设置窗口取消勾选

## 验证

- Release 构建通过，0 警告、0 错误
- 本地 release package 安装注册通过
- PowerPoint 实机自动化验证通过：新建和编辑在开关开启时显示弹窗，关闭时不显示

## 相关文档

- [v1.0.2 E2E 测试报告](./e2e-test-report-v1.0.2.md)
- [v1.0.2 发布证据与日志索引](./release-evidence-v1.0.2.md)
