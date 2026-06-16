# Full E2E Test Report (v1.0.4)

[English](./e2e-test-report-v1.0.4.en.md) | [中文](./e2e-test-report-v1.0.4.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This report records the real release validation results for `DrawioPpt v1.0.4`, focusing on release package generation, the unified Office registration script, the PowerPoint/Word installation flow, and Word COM Add-in loading.

## 1. Test Environment

- Operating system: Windows
- Office hosts: Microsoft PowerPoint Desktop and Microsoft Word Desktop
- Build configuration: `Release|x64`
- Registration mode: current-user COM Add-in registration
- Install root: temporary directory `%TEMP%\DrawioPptReleaseV104-<guid>`

## 2. Commands

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.4 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.4\package\scripts\install-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV104-<guid>
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.4\package\scripts\uninstall-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV104-<guid>
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Release
```

PowerShell AST parse checks and package content checks were also run for the new unified register/unregister scripts.

## 3. Build and Package Results

- Release build passed.
- Build result: 0 warnings, 0 errors.
- Release folder generated: `artifacts\releases\v1.0.4\package`
- Release zip generated: `artifacts\releases\v1.0.4\DrawioPpt-v1.0.4.zip`
- Package content inspection confirmed no PDB, log, or screenshot assets.

## 4. Installation Validation

- The release installer successfully called `scripts\register-office-addins.ps1`.
- The PowerPoint COM Add-in registration entry was written: `Greensoft.DrawioPptAddIn`.
- The Word COM Add-in registration entry was written: `Greensoft.DrawioWordAddIn`.
- Both PowerPoint and Word registration entries had `LoadBehavior=3`.
- The release uninstaller successfully called `scripts\unregister-office-addins.ps1`.
- After uninstall, both PowerPoint and Word registration entries were removed.

After temporary install validation, the local default Office add-in registration was restored.

## 5. Word Load Validation

`scripts\word-addin-load-check.ps1 -Configuration Release` result:

```text
WordAddInConnect=True
```

This confirms that Word can load and connect the Release build of the Word COM Add-in.

## 6. Known Limitations

This release focuses on installer scripts and package manifests. It does not change the Draw.io editor protocol, SVG generation logic, or Word picture write-back logic.
