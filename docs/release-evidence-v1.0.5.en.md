# v1.0.5 Release Evidence and Log Index

[English](./release-evidence-v1.0.5.en.md) | [中文](./release-evidence-v1.0.5.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This document summarizes the deliverables, real validation results, and release information for `DrawioPpt v1.0.5`.

## 1. Deliverables

- Version: `v1.0.5`
- Release folder: `artifacts\releases\v1.0.5\package`
- Release zip: `artifacts\releases\v1.0.5\DrawioPpt-v1.0.5.zip`
- Install entry: `install.cmd`
- Unified register script: `scripts\register-office-addins.ps1`
- Unified unregister script: `scripts\unregister-office-addins.ps1`
- Install acceptance script: `scripts\verify-office-install.ps1`
- PowerPoint add-in: `bin\DrawioPpt.PowerPointAddIn.dll`
- Word add-in: `bin\DrawioPpt.WordAddIn.dll`

## 2. Documentation

- [v1.0.5 release notes](./release-notes-v1.0.5.en.md)
- [v1.0.5 E2E test report](./e2e-test-report-v1.0.5.en.md)
- [User guide](./user-guide.en.md)
- [Installation and local debugging guide](./installation.en.md)
- [Word add-in design and usage](./word-addin.en.md)
- [Architecture](./architecture.en.md)

## 3. Validation Summary

- Reproduced the default install directory state where Word was not registered.
- Downloaded and retested the `v1.0.4` GitHub asset; it included the unified registration script and could install Word to the default path.
- Reproduced that `word-addin-load-check.ps1` removed Word registration.
- Retested after fixing `word-addin-load-check.ps1`; Word registration and `CodeBase` were preserved.
- New install-acceptance coverage test failed first, then passed.
- Default install directory acceptance passed; both PowerPoint and Word registration and COM loading passed.
- Word disabled-items check found no DrawioWord disabled entry.
- `v1.0.5` release build, packaging, default installation, and acceptance are release gates.

## 4. Distribution Notes

- For external distribution, provide `DrawioPpt-v1.0.5.zip` first.
- After extraction, users run `install.cmd`, which registers both PowerPoint and Word and runs non-interactive install acceptance checks.
- If Word still does not appear after installation, run `scripts\verify-office-install.ps1` first; it distinguishes missing registration, wrong `CodeBase`, Office disabled items, and COM load failures.
