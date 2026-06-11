# v1.0.3 Release Evidence and Log Index

[English](./release-evidence-v1.0.3.en.md) | [中文](./release-evidence-v1.0.3.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This document summarizes the deliverables, real validation results, and release information for `DrawioPpt v1.0.3`.

## 1. Deliverables

- Version: `v1.0.3`
- Release folder: `artifacts\releases\v1.0.3\package`
- Release zip: `artifacts\releases\v1.0.3\DrawioPpt-v1.0.3.zip`
- Install entry point: `install.cmd`
- PowerShell installer: `scripts\install-release.ps1`
- Uninstall entry point: `uninstall.cmd`
- PowerPoint add-in: `bin\DrawioPpt.PowerPointAddIn.dll`
- Word add-in: `bin\DrawioPpt.WordAddIn.dll`

## 2. Docs

- [v1.0.3 release notes](./release-notes-v1.0.3.en.md)
- [v1.0.3 E2E test report](./e2e-test-report-v1.0.3.en.md)
- [User guide](./user-guide.en.md)
- [Installation and local debugging guide](./installation.en.md)
- [Word add-in design and usage](./word-addin.en.md)

## 3. Validation Summary

- Release build: passed with 0 warnings and 0 errors.
- Release packaging: passed.
- Package inspection: no PDB, log, or screenshot assets.
- Temporary installation: passed, both PowerPoint and Word COM Add-ins registered successfully.
- Temporary uninstall: passed, both PowerPoint and Word COM Add-ins unregistered successfully.

## 4. Recommended External Editor

- Name: `draw.io Desktop v30.0.4 XML tools build`
- Release: https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1
- Asset: `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip`
- SHA256: `B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D`
- Runtime: fully extract and run `win-unpacked\draw.io.exe`

## 5. Distribution Notes

- For external distribution, provide `DrawioPpt-v1.0.3.zip` first.
- Users can extract the package and run `install.cmd`; it registers both PowerPoint and Word add-ins.
- If an older version is installed, close PowerPoint/Word and rerun the installer to upgrade.
