# v1.0.2 Release Evidence and Log Index

[English](./release-evidence-v1.0.2.en.md) | [中文](./release-evidence-v1.0.2.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This document summarizes the deliverables, real validation results, and log locations for `DrawioPpt v1.0.2`.

## 1. Deliverables

- Version: `v1.0.2`
- Release folder: `artifacts\releases\v1.0.2\package`
- Release zip: `artifacts\releases\v1.0.2\DrawioPpt-v1.0.2.zip`
- Install entry point: `install.cmd`
- PowerShell installer: `scripts\install-release.ps1`
- Uninstall entry point: `uninstall.cmd`

## 2. Docs

- [v1.0.2 release notes](./release-notes-v1.0.2.en.md)
- [v1.0.2 E2E test report](./e2e-test-report-v1.0.2.en.md)
- [User guide](./user-guide.en.md)
- [Architecture](./architecture.en.md)

## 3. Validation Summary

- Release build: passed with 0 warnings and 0 errors
- Local installation: passed
- PowerPoint COM Add-in registration: `Greensoft.DrawioPptAddIn`, `LoadBehavior=3`
- PowerPoint create/edit setting toggle test: passed

## 4. Log Location

```text
C:\Users\nabul\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

The test log records create, edit, working-file preparation, and desktop editor launch paths from this validation run.

## 5. Distribution Notes

- For external distribution, provide `DrawioPpt-v1.0.2.zip` first
- Users can extract the package and run `install.cmd` to install the current-user PowerPoint add-in
- If an older version is installed, rerunning the installer is enough to move to `v1.0.2`
