# v1.0.10 发布证据

本文件记录 v1.0.10 发布包的构建标识、压缩包 SHA-256、完整 Office E2E 和标准本地安装验证结果。

- Release x64 构建通过，`0` 个警告、`0` 个错误。
- `install-release-encoding-safety-test.ps1` 通过：仓库和发布包内的全部 `.ps1` 都是 ASCII，且 `install.cmd` 会输出失败退出码并返回该退出码。
- `install-release-upgrade-safety-test.ps1` 通过：覆盖同目录和子目录升级。
- 完整发布包 Office E2E：`14/14 PASS`，清理成功。
- 临时发布包安装中，Word、PowerPoint 均完成真实 COM 加载，并返回与安装包 `BuildInfo.txt` 一致的 BuildId。

发布压缩包的最终 BuildId 和 SHA-256 以最终包根目录的 `BuildInfo.txt`、`PACKAGE.txt` 及 GitHub Release digest 为准。
