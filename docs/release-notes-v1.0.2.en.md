# v1.0.2

[English](./release-notes-v1.0.2.en.md) | [中文](./release-notes-v1.0.2.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.2` is a small editing-experience and delivery release. It adds a setting that controls whether diagram information is shown when creating or editing Draw.io diagrams, and it ships as an installable release package.

## Problems Solved First

- Diagram information popups were useful for troubleshooting, but interrupted experienced users during routine create/edit work.
- Users still needed a way to keep those details available when checking diagram IDs, working-file paths, or editor mode.
- Users needed the option to suppress ordinary success prompts while still seeing real error messages.
- The release needed to look more like an end-user installable package instead of relying only on repository registration scripts.

## Changes

- Added `Show diagram information when creating or editing` to the settings dialog
- Added the internal `ShowDiagramInfoDialog` setting, defaulting to enabled for older settings files
- Unified successful diagram information dialogs for create and edit flows under this setting
- Kept error prompts independent from this setting so failures remain visible
- Updated README, architecture docs, and user guides

## Install and Upgrade

- The release package includes `install.cmd`, `uninstall.cmd`, and `scripts\install-release.ps1`
- On machines with an older registered add-in, rerunning the installer from the package moves PowerPoint to `v1.0.2`
- The default configuration still shows diagram information; disable it from the settings dialog if you prefer a quieter workflow

## Validation

- Release build passed with 0 warnings and 0 errors
- Local release package installation and registration passed
- PowerPoint host automation passed for both create and edit flows: dialogs are shown when enabled and suppressed when disabled

## Related Docs

- [v1.0.2 E2E test report](./e2e-test-report-v1.0.2.en.md)
- [v1.0.2 release evidence and log index](./release-evidence-v1.0.2.en.md)
