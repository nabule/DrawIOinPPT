# v1.0.1 Release Evidence and Log Index

[English](./release-evidence-v1.0.1.en.md) | [中文](./release-evidence-v1.0.1.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This document summarizes the deliverables, real validation results, and log locations for `DrawioPpt v1.0.1`, with emphasis on the SVG text vector-fidelity fix delivered in this version.

## Release Information

- Version: `v1.0.1`
- Release date: `2026-04-20`
- Assembly versions:
  - `DrawioPpt.Core = 1.0.1.0`
  - `DrawioPpt.PowerPointAddIn = 1.0.1.0`

## Program Deliverables

- Release directory: `artifacts\releases\v1.0.1\package`
- Release zip: `artifacts\releases\v1.0.1\DrawioPpt-v1.0.1.zip`

## Documentation Deliverables

- [v1.0.1 release notes](./release-notes-v1.0.1.en.md)
- [v1.0.1 E2E test report](./e2e-test-report-v1.0.1.en.md)
- [Installation and local debugging guide](./installation.en.md)
- [User guide](./user-guide.en.md)

## Real Validation Results

- Release build: passed
  - `MSBuild Rebuild (Release|x64)`
  - `0 warning / 0 error`
  - `DrawioPpt.Core.dll = 1.0.1.0`
  - `DrawioPpt.PowerPointAddIn.dll = 1.0.1.0`
- Release-package installation verification: passed
  - Install root: `C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1`
  - Registry `CodeBase`: `file:///C:/Users/nabul/AppData/Local/Temp/DrawioPpt/release-verify-v1.0.1/bin/DrawioPpt.PowerPointAddIn.dll`
- Packaged URL smoke: passed
  - The packaged `url-editor-smoke.ps1 -SkipBuild` returned `0`
  - The plug-in log recorded `configure -> init -> load -> save -> export -> DialogResult=OK`
- SVG vector-text regression:
  - `OriginalForeignObject=True`
  - `OriginalEmbeddedTextPng=True`
  - `SanitizedForeignObject=False`
  - `SanitizedEmbeddedTextPng=False`
  - `StoredSvgContainsForeignObject=False`
  - `StoredSvgContainsEmbeddedTextPng=False`
  - `StoredSvgHasTspan=True`

## Log Index

```text
artifacts\logs\v1.0.1\
C:\Users\nabul\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

## Delivery Recommendations

- For external distribution, provide `DrawioPpt-v1.0.1.zip` first
- If acceptance evidence is needed, also provide `docs\release-evidence-v1.0.1.en.md`, `docs\e2e-test-report-v1.0.1.en.md`, and the `logs\` directory
