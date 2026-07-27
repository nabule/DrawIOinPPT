# v1.0.6 Release Evidence and Log Index

[English](./release-evidence-v1.0.6.en.md) | [中文](./release-evidence-v1.0.6.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Deliverables

- Version: `v1.0.6`
- Release folder: `artifacts\releases\v1.0.6\package`
- Release archive: `artifacts\releases\v1.0.6\DrawioPpt-v1.0.6.zip`
- New release-verification script: `scripts\word-selection-event-e2e.ps1`

## Release Gates

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.6 -Configuration Release -Platform x64 -SkipBuild
```

The release package must contain `DrawioPpt.WordAddIn.dll`, `word-selection-event-e2e.ps1`, the versioned bilingual release notes, the E2E report, and this evidence index. Validate the release DLL with:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.6\package\scripts\word-selection-event-e2e.ps1 -SkipBuild -SkipSourceCheck -AssemblyRoot .\artifacts\releases\v1.0.6\package\bin
```

Then run `verify-office-install.ps1`, `word-addin-load-check.ps1`, and the full Office E2E against a temporary installation. Their output and the GitHub Release asset SHA256 are part of the release record.

## Final Verification Results

- The `Release|x64` build passed with 0 warnings and 0 errors.
- Both the source and release-DLL `word-selection-event-e2e.ps1` runs passed with `NoSelectionPolling=True`, `ManagedSelectionDetected=True`, and `PlainPictureCanBind=True`.
- `word-url-e2e.ps1` passed: its mock HTTP server was reachable, creation, reopen-edit, and final reopen persistence were all `True`, and the add-in setting was restored to `LoadBehavior=3`.
- The release package was installed into a temporary directory; `verify-office-install.ps1` verified Word/PowerPoint registration and live COM loading, and `word-addin-load-check.ps1` reported `WordAddInConnect=True`.
- The full Office E2E report passed 7/7 checks: release package, install, Office verification, PowerPoint loading, installed URL smoke, PowerPoint URL E2E, and desktop draw.io export. Office-process cleanup also passed.

## Test Preconditions

Word COM E2E requires Word to be closed and must not run alongside another Word automation task. To isolate its host, `word-url-e2e.ps1` temporarily sets the current-user Word add-in `LoadBehavior` to `0` and restores the original value in `finally`; do not start Word while it is running.
