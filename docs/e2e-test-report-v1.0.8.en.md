# v1.0.8 E2E Test Report

[English](./e2e-test-report-v1.0.8.en.md) | [中文](./e2e-test-report-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Test Environment

- Date: 2026-07-28
- Windows + Microsoft Word Desktop / PowerPoint Desktop
- .NET Framework 4.8, `Release|x64`
- Current-user Office COM Add-in registration and a temporarily installed v1.0.8 release package

## Pre-release Gates

The commands below ran serially. No user Word or PowerPoint process was present before they started, and no `WINWORD` or `POWERPNT` process remained afterward.

| Check | Result | Evidence |
| --- | --- | --- |
| Release x64 build | PASS | 0 warnings, 0 errors |
| Word selection-sync source contract | PASS | `WORD_SELECTION_SYNC_TEST_PASS` |
| WebView2 user-data-folder source contract | PASS | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Word URL-host safety contract | PASS | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| Full-E2E cleanup safety contract | PASS | `FULL_E2E_CLEANUP_SAFETY_TEST_PASS`; a mismatched start time protects the process, cleanup requires matching PID and start time, and a non-zero cleanup command propagates to final failure |
| Word complex-metadata E2E | PASS | `DrawioXmlChars=271361`, `AlternativeTextChars=436`; primary storage, lightweight reference, failure fallback, legacy migration, stable part ID, and save-close-reopen persistence all satisfied their assertions |
| Word selection-event E2E | PASS | `NoSelectionPolling=True`, `ManagedSelectionDetected=True`, `PlainPictureCanBind=True` |
| Word URL E2E | PASS | `CreatedManagedPicture=True`, `ReopenedEditApplied=True`, `PersistedAfterReopen=True` |
| Real Word URL-host E2E | PASS | `WordAddInConnect=True`, `ActualHostProcess=WINWORD`, `ActualWordUrlEditorSaved=True`, `WebView2UserDataFolderExists=True`, `AccessDeniedDialog=False` |

The complex-metadata regression also reported `StoredPayload=True`, `LightweightAlternativeText=True`, `FallbackRetainsPayload=True`, `LegacyMigrationPassed=True`, `LegacyMigrationPartIdsStable=True`, `FallbackPersistedAfterReopen=True`, `LegacyMigrationPersistedAfterReopen=True`, and `CleanupResidualWinWord=False`.

## Full Office E2E Against the Temporary Installation

Command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

Result: 9/9 passed.

Packaging first reported `PackageMarkdownMissingLinkCount=0`, confirming that every local Markdown link in the archive resolves to a packaged file.

| Check | Result | Evidence |
| --- | --- | --- |
| `ReleasePackage` | PASS | The v1.0.8 release directory was created |
| `InstallRelease` | PASS | The package was installed into an isolated temporary directory |
| `InstalledOfficeAddInsVerify` | PASS | Word / PowerPoint registration and COM loading passed |
| `PowerPointAddInLoad` | PASS | `Connect=True` |
| `InstalledUrlSmoke` | PASS | `Saved=True`, `SvgHasSmoke=True`, `XmlHasSmokeId=True` |
| `InstalledWordUrlHostE2E` | PASS | DLLs from the temporary package completed URL write-back in real `WINWORD.EXE` |
| `InstalledWordComplexMetadataE2E` | PASS | DLLs from the temporary package passed the complex-XML primary-store, lightweight-reference, fallback, and migration regression |
| `PowerPointUrlE2E` | PASS | `CreatedManagedShape=True`, `ReopenedEditApplied=True`, `PersistedAfterReopen=True`, `ExitCode=0` |
| `DesktopExporter` | PASS | draw.io Desktop exported SVG successfully |

The PowerPoint mock editor obtains a preferred loopback port from the operating system and uses C# `StartWithRetry` to handle a bind race. Reachability checks and `PluginSettings.EditorUrl` both use the final bound port. This run reported `MockServerPort=3721`.

The installed complex-metadata test wrote `TestWordProcessIdentity=4132:639208146977213246`. The finally block checked the test-owned Word process by PID and start time; it had already exited normally, so the run reported `TestOwnedProcessAlreadyExited=4132`. Temporary-package uninstall, repository add-in registration restoration, and settings restoration completed, followed by `FullE2ECleanupSucceeded=True`.

## Reports and Logs

- Summary report: `artifacts\test-reports\full-e2e-v1.0.8.md`
- Report copy: `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- PowerShell transcript: `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- Plug-in log snapshot: `artifacts\logs\v1.0.8\drawioppt-full-e2e.log`
- Package contents and DLL hashes: [v1.0.8 release evidence](./release-evidence-v1.0.8.en.md)

## Acceptance Boundary

This report verifies the source Release DLLs and a temporary package installation. Final installation into the standard local directory and the manual 10-second drag, resize, reposition, and reopen/edit comparison remain a later acceptance step and are not claimed here.
