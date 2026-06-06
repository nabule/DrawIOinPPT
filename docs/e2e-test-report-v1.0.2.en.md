# Full E2E Test Report (v1.0.2)

[English](./e2e-test-report-v1.0.2.en.md) | [中文](./e2e-test-report-v1.0.2.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This report records the real release validation results for `DrawioPpt v1.0.2`, focusing on the new diagram information dialog setting, installable package deployment, and PowerPoint host behavior.

## 1. Environment

- OS: Windows
- PowerPoint: Microsoft PowerPoint Desktop
- Build configuration: `Release|x64`
- Install root: `C:\Users\nabul\AppData\Local\Greensoft\DrawioPpt`
- Registration: current-user PowerPoint COM Add-in registration

## 2. Commands

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.2 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.2\package\scripts\install-release.ps1
```

## 3. Build and Install Results

- Release build passed
- Build result: 0 warnings and 0 errors
- Release folder generated: `artifacts\releases\v1.0.2\package`
- Release zip generated: `artifacts\releases\v1.0.2\DrawioPpt-v1.0.2.zip`
- Installer successfully registered `Greensoft.DrawioPptAddIn`
- Registry `LoadBehavior=3`

## 4. PowerPoint Host Validation

Real PowerPoint host automation was used to start PowerPoint and call the add-in's main flows.

```text
Case   ShowDialogSetting ShapesBefore ShapesAfter ClosedDialogs
Create True              0            1           1
Create False             0            1           0
Edit   True              0            1           1
Edit   False             0            1           0
```

Conclusion:

- Creating a Draw.io diagram shows the information dialog when enabled and suppresses it when disabled
- Editing an existing Draw.io diagram shows the information dialog when enabled and suppresses it when disabled
- All four cases inserted or recognized an add-in-managed shape successfully
- The local machine configuration was restored after testing: desktop editor path is `C:\Program Files\draw.io\draw.io.exe`

## 5. Known Limitation

`scripts\url-editor-smoke.ps1` previously timed out in the current automation environment because the WinForms/WebView2 window did not exit automatically. This release does not change the URL editor save protocol, so validation focused on the new setting and the desktop create/edit entry points.
