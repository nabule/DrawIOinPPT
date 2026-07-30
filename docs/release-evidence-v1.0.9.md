# v1.0.9 发布证据

本文件记录 v1.0.9 发布包的构建标识、压缩包 SHA-256、完整 Office E2E 结果和标准安装验证结果。

- Release x64 构建通过，`0` 个警告、`0` 个错误。
- `version-display-safety-test.ps1` 通过，覆盖 Word/PPT Ribbon、COM 入口、自动化版本接口及安装验收。
- `install-release-upgrade-safety-test.ps1` 通过，覆盖同目录和子目录升级。
- 完整发布包 Office E2E：`13/13 PASS`，清理成功。
- 标准本地安装：Word、PowerPoint 均完成真实 COM 加载，并返回与安装包 `BuildInfo.txt` 一致的 BuildId。

发布压缩包和实际 BuildId 以最终包根目录的 `BuildInfo.txt`、`PACKAGE.txt` 及 GitHub Release SHA-256 digest 为准。
