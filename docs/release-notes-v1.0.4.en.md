# v1.0.4

[English](./release-notes-v1.0.4.en.md) | [中文](./release-notes-v1.0.4.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.4` is an installer-focused release. It adds a unified repository and release-package registration entry so the PowerPoint and Word add-ins can be registered and unregistered together.

## Highlights

- Adds `scripts\register-office-addins.ps1` to register both current-user COM Add-ins: PowerPoint and Word.
- Adds `scripts\unregister-office-addins.ps1` to unregister both COM Add-ins together.
- Changes release `install.cmd` / `scripts\install-release.ps1` to call the unified registration entry, preventing install flows that only cover PowerPoint.
- Changes release `uninstall.cmd` / `scripts\uninstall-release.ps1` to call the unified unregister entry.
- Adds the unified scripts to the release package and `PACKAGE.txt` manifest.
- Updates the README, installation guide, user guide, architecture docs, and regression checklist with the dual-host installation flow.

## Installation

End users should download from the Release page:

```text
DrawioPpt-v1.0.4.zip
```

After extraction, close PowerPoint and Word, then run:

```powershell
.\install.cmd
```

The installer registers both:

- `Greensoft.DrawioPptAddIn`
- `Greensoft.DrawioWordAddIn`

For repository-based development, build first and then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-office-addins.ps1
```

For single-host diagnostics, `scripts\register-addin.ps1` and `scripts\register-word-addin.ps1` remain available.

## Recommended draw.io Desktop Version

Desktop mode still works best with [draw.io Desktop v30.0.4 XML tools build](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1).

Recommended asset:

```text
draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip
```

After extraction, set `Desktop Path` to:

```text
win-unpacked\draw.io.exe
```

This build can copy and paste Draw.io XML directly from the toolbar, which helps diagnose source data embedded in Office files.

## Validation

- PowerShell parse checks passed.
- Release build passed with 0 warnings and 0 errors.
- Release package generation passed.
- Package content inspection confirmed PowerPoint/Word add-ins and the unified register/unregister scripts are included.
- Temporary package install and uninstall passed; both PowerPoint and Word registration entries were written and removed.
- Word COM Add-in load check passed with `WordAddInConnect=True`.

Details:

- [v1.0.4 E2E test report](./e2e-test-report-v1.0.4.en.md)
- [v1.0.4 release evidence and log index](./release-evidence-v1.0.4.en.md)
