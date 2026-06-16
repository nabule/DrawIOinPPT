# v1.0.4 Release Evidence and Log Index

[English](./release-evidence-v1.0.4.en.md) | [中文](./release-evidence-v1.0.4.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This document summarizes the deliverables, real validation results, and release information for `DrawioPpt v1.0.4`.

## 1. Deliverables

- Version: `v1.0.4`
- Release folder: `artifacts\releases\v1.0.4\package`
- Release zip: `artifacts\releases\v1.0.4\DrawioPpt-v1.0.4.zip`
- Install entry: `install.cmd`
- Unified register script: `scripts\register-office-addins.ps1`
- Unified unregister script: `scripts\unregister-office-addins.ps1`
- PowerPoint add-in: `bin\DrawioPpt.PowerPointAddIn.dll`
- Word add-in: `bin\DrawioPpt.WordAddIn.dll`

## 2. Documentation

- [v1.0.4 release notes](./release-notes-v1.0.4.en.md)
- [v1.0.4 E2E test report](./e2e-test-report-v1.0.4.en.md)
- [User guide](./user-guide.en.md)
- [Installation and local debugging guide](./installation.en.md)
- [Word add-in design and usage](./word-addin.en.md)
- [Architecture](./architecture.en.md)

## 3. Validation Summary

- New unified installer-entry coverage test: failed first, then passed.
- PowerShell parse checks: passed.
- Release build: passed with 0 warnings and 0 errors.
- Release packaging: passed.
- Package content check: unified register/unregister scripts included; no PDB, log, or screenshot assets included.
- Temporary install: passed; both PowerPoint and Word COM Add-ins were registered.
- Temporary uninstall: passed; both PowerPoint and Word COM Add-ins were unregistered.
- Word load check: passed with `WordAddInConnect=True`.

## 4. Recommended External Editor

- Name: `draw.io Desktop v30.0.4 XML tools build`
- Release: https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1
- Asset: `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip`
- SHA256: `B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D`
- Run mode: fully extract the archive and run `win-unpacked\draw.io.exe`

## 5. Distribution Notes

- For external distribution, provide `DrawioPpt-v1.0.4.zip` first.
- After extraction, users run `install.cmd`, which registers both the PowerPoint and Word add-ins.
- To upgrade an existing install, close PowerPoint/Word and run the installer again.
- Repository-based development uses `scripts\register-office-addins.ps1` to register both host add-ins.
