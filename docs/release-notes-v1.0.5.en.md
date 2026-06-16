# v1.0.5

[English](./release-notes-v1.0.5.en.md) | [中文](./release-notes-v1.0.5.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.5` is a Word installation acceptance fix. The `v1.0.4` release package already included the Word registration entry, but the old `word-addin-load-check.ps1` unregistered the Word add-in at the end of the test, leaving the default environment in a "Word is not installed" state after the check. This release fixes that test script and adds default-install-path plus real Word COM loading acceptance.

## Highlights

- Adds `scripts\verify-office-install.ps1` to validate release-package installation results.
- Fixes `scripts\word-addin-load-check.ps1`: it now records any existing Word registration before the test and restores it afterward instead of uninstalling the default Word add-in registration.
- `install-release.ps1` now runs non-interactive acceptance checks after registering PowerPoint and Word, confirming both host DLLs, COM registration, and `CodeBase` targets.
- The full E2E script now runs post-install acceptance and actually creates Word/PowerPoint COM applications to connect `COMAddIns`.
- The release package manifest now includes `scripts\verify-office-install.ps1`.
- README, installation guide, architecture docs, and user guide are updated for `v1.0.5` with the post-install acceptance command.

## Root Cause

The re-check found that the default install directory had stale package contents at one point and the current-user PowerPoint/Word registration entries were missing. Downloading the `v1.0.4` asset from GitHub and running `install.cmd` registered and loaded Word successfully, so the online asset did include the Word registration scripts.

Further reproduction confirmed that the old `word-addin-load-check.ps1` called `unregister-word-addin.ps1` in `finally`. Running the check script therefore removed Word registration. `v1.0.5` restores the pre-check registration state and makes real Word COM loading from the default install path an explicit release gate.

## Post-Install Acceptance

After installation, run:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

Expected output includes:

```text
Word registration verified: Greensoft.DrawioWordAddIn
Word.Application COM load verified: Greensoft.DrawioWordAddIn
PowerPoint.Application COM load verified: Greensoft.DrawioPptAddIn
DrawioPpt Office installation verified.
```

## Validation

- `word-addin-load-check.ps1` registration-preservation regression check failed first, then passed.
- New install-acceptance coverage check failed first, then passed.
- Default install directory acceptance passed; Word and PowerPoint both loaded through Office COM.
- The `v1.0.4` GitHub Release asset was downloaded and retested, confirming it included the unified registration scripts and that Word loaded after default installation.
- `v1.0.5` release build, packaging, default install, install acceptance, and Word load checks are all release gates.

Details:

- [v1.0.5 E2E test report](./e2e-test-report-v1.0.5.en.md)
- [v1.0.5 release evidence and log index](./release-evidence-v1.0.5.en.md)
