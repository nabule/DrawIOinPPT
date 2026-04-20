# v1.0.1 发布证据与日志索引

[中文](./release-evidence-v1.0.1.md) | [English](./release-evidence-v1.0.1.en.md) | [中文首页](../README.md) | [English Home](../README.en.md)

本文档汇总 `DrawioPpt v1.0.1` 的程序交付物、真实验证结果和日志位置，重点对应本版本的 SVG 文字矢量保真修复。

## 发布信息

- 版本：`v1.0.1`
- 发布日期：`2026-04-20`
- 程序集版本：
  - `DrawioPpt.Core = 1.0.1.0`
  - `DrawioPpt.PowerPointAddIn = 1.0.1.0`

## 程序交付物

- 发布目录：`artifacts\releases\v1.0.1\package`
- 发布压缩包：`artifacts\releases\v1.0.1\DrawioPpt-v1.0.1.zip`

## 文档交付物

- [v1.0.1 发布说明](./release-notes-v1.0.1.md)
- [v1.0.1 E2E 测试报告](./e2e-test-report-v1.0.1.md)
- [安装与联调说明](./installation.md)
- [用户手册](./user-guide.md)

## 实际测试情况

- Release 构建：通过
  - `MSBuild Rebuild (Release|x64)`
  - `0 warning / 0 error`
  - `DrawioPpt.Core.dll = 1.0.1.0`
  - `DrawioPpt.PowerPointAddIn.dll = 1.0.1.0`
- 发布包安装验证：通过
  - 安装目录：`C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1`
  - 注册表 `CodeBase`：`file:///C:/Users/nabul/AppData/Local/Temp/DrawioPpt/release-verify-v1.0.1/bin/DrawioPpt.PowerPointAddIn.dll`
- 安装包 URL smoke：通过
  - 安装包内 `url-editor-smoke.ps1 -SkipBuild` 返回 `0`
  - 插件日志记录了 `configure -> init -> load -> save -> export -> DialogResult=OK`
- SVG 矢量文字回归：
  - `OriginalForeignObject=True`
  - `OriginalEmbeddedTextPng=True`
  - `SanitizedForeignObject=False`
  - `SanitizedEmbeddedTextPng=False`
  - `StoredSvgContainsForeignObject=False`
  - `StoredSvgContainsEmbeddedTextPng=False`
  - `StoredSvgHasTspan=True`

## 日志索引

```text
artifacts\logs\v1.0.1\
C:\Users\nabul\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

## 交付建议

- 对外分发时优先提供 `DrawioPpt-v1.0.1.zip`
- 若需要验收依据，连同 `docs\release-evidence-v1.0.1.md`、`docs\e2e-test-report-v1.0.1.md` 和 `logs\` 目录一起提供
