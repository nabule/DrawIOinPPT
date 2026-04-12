# v1.0.0 Release Evidence and Log Index

[English](./release-evidence-v1.0.0.en.md) | [中文](./release-evidence-v1.0.0.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This document summarizes the program deliverables, real validation results, and log locations for `DrawioPpt v1.0.0`, making release, acceptance, and troubleshooting easier to trace.

## Release Information

- Version: `v1.0.0`
- Release date: `2026-04-12`
- Release commit: `6239eab246aa603a90af119b640aaff1186f95dc`
- Assembly versions:
  - `DrawioPpt.Core = 1.0.0.0`
  - `DrawioPpt.PowerPointAddIn = 1.0.0.0`

## Program Deliverables

- Release directory: `artifacts\releases\v1.0.0\package`
- Release zip: `artifacts\releases\v1.0.0\DrawioPpt-v1.0.0.zip`
- Main binaries:
  - `bin\DrawioPpt.PowerPointAddIn.dll`
  - `bin\DrawioPpt.Core.dll`
  - `bin\Microsoft.Web.WebView2.Core.dll`
  - `bin\Microsoft.Web.WebView2.WinForms.dll`
  - `bin\runtimes\win-x64\native\WebView2Loader.dll`
- Install entry points:
  - `install.cmd`
  - `scripts\install-release.ps1`
- Uninstall entry points:
  - `uninstall.cmd`
  - `scripts\uninstall-release.ps1`

## Documentation Deliverables

- [Installation and local debugging guide](./installation.en.md)
- [User guide](./user-guide.en.md)
- [Regression checklist](./regression-checklist.en.md)
- [v1.0.0 release notes](./release-notes-v1.0.0.en.md)
- [v1.0.0 E2E test report](./e2e-test-report-v1.0.0.en.md)

## Real Validation Results

Execution date: `2026-04-12`

Validation environment:

- OS: `Microsoft Windows 10 Pro 10.0.19045 64-bit`
- PowerPoint: `Microsoft PowerPoint Desktop x64`
- draw.io Desktop: `C:\Program Files\draw.io\draw.io.exe`
- WebView2 Runtime: installed

Commands actually executed:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1 -SkipBuild
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.0 -SkipBuild
```

Validation conclusions:

- Environment check passed
- Release build passed with `0` warnings and `0` errors
- URL smoke test passed
- Release packaging, installation, registration, and PowerPoint host loading passed
- URL-mode create, save, close/reopen, re-edit, and persistence checks after reopening passed
- draw.io Desktop CLI export passed

Key result excerpts:

- `CreatedManagedShape=True`
- `ReopenedEditApplied=True`
- `PersistedAfterReopen=True`
- `FinalCustomXmlCount=4`
- `ShapeCount=1`

These results indicate that URL-mode write-back, second-round editing, and `CustomXMLParts` persistence have all been validated on a real PowerPoint host for `v1.0.0` on the current machine.

## Log Index

All release-related logs for this version are stored in:

```text
artifacts\logs\v1.0.0\
```

Key log files:

- `dev-check.log`
  Environment-check output
- `build-release.log`
  Release-build output
- `url-editor-smoke.log`
  URL-mode smoke-test console output
- `drawioppt-url-editor-smoke.log`
  Add-in log snapshot captured during the URL smoke test
- `full-e2e.console.log`
  Full E2E console output
- `full-e2e-v1.0.0.transcript.log`
  Full E2E PowerShell transcript
- `full-e2e-v1.0.0.report.md`
  Raw E2E report copy
- `drawioppt-full-e2e.log`
  Add-in log snapshot captured during the full E2E run

## Delivery Recommendations

- For external distribution, provide `DrawioPpt-v1.0.0.zip` first
- If acceptance evidence is needed, also provide `docs\release-evidence-v1.0.0.en.md`, `docs\e2e-test-report-v1.0.0.en.md`, and the `logs\` directory
- For on-site troubleshooting, prioritize `logs\drawioppt-full-e2e.log` and the user's local `%APPDATA%\Greensoft\DrawioPpt\Logs\drawioppt.log`
