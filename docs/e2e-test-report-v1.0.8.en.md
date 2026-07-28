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
| Word zero-passive-selection-path source contract | PASS | `WORD_SELECTION_SYNC_TEST_PASS`: no `WindowSelectionChange`, auto-open, or startup selection read; each of the four explicit commands captures one live picture snapshot |
| WebView2 user-data-folder source contract | PASS | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Word URL-host safety contract | PASS | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| Full-E2E cleanup and unified-exit safety contract | PASS | `FULL_E2E_CLEANUP_SAFETY_TEST_PASS`; the main script AST contains only the final `exit 1`, a mismatched start time protects the process, and a controlled probe writes the main, uninstall, re-registration, and settings-restoration errors to one FatalError row before returning 1 |
| Word complex-metadata E2E | PASS | `DrawioXmlChars=271361`, `AlternativeTextChars=436`; primary storage, lightweight reference, failure fallback, legacy migration, stable part ID, and save-close-reopen persistence all satisfied their assertions |
| Word explicit-selection recognition E2E | PASS | `NoPassiveSelectionMetadataPath=True`, `ExplicitManagedSelectionDetected=True`, `ExplicitPlainPictureCanBind=True` |
| Word URL E2E | PASS | `CreatedManagedPicture=True`, `ReopenedEditApplied=True`, `PersistedAfterReopen=True` |
| Real Word URL-host E2E | PASS | `WordAddInConnect=True`, `ActualHostProcess=WINWORD`, `ExplicitEditAutomationAvailable=True`, `ExplicitEditCommandInvoked=True`, `ActualWordUrlEditorSaved=True`, `AccessDeniedDialog=False` |

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
| `InstalledWordUrlHostE2E` | PASS | DLLs from the temporary package completed URL write-back through the select-then-explicit-command flow in real `WINWORD.EXE` |
| `InstalledWordComplexMetadataE2E` | PASS | DLLs from the temporary package passed the complex-XML primary-store, lightweight-reference, fallback, and migration regression |
| `PowerPointUrlE2E` | PASS | `CreatedManagedShape=True`, `ReopenedEditApplied=True`, `PersistedAfterReopen=True`, `ExitCode=0` |
| `DesktopExporter` | PASS | draw.io Desktop exported SVG successfully |

The URL-smoke mock editor tries consecutive ports when the preferred loopback port is excluded or occupied. Reachability checks and `PluginSettings.EditorUrl` both use the final bound port. This run reported `MockServerPort=8752` for smoke and `MockServerPort=12440` for the real Word URL host.

The installed complex-metadata test wrote `TestWordProcessIdentity=42848:639208283590231895`. The finally block checked the test-owned Word process by PID and start time; it had already exited normally, so the run reported `TestOwnedProcessAlreadyExited=42848`. Temporary-package uninstall, repository add-in registration restoration, and settings restoration completed, followed by `FullE2ECleanupSucceeded=True`.

## Subsequent Standard Local Installation and Word UI-thread Stress Confirmation

- Standard install root: `%LOCALAPPDATA%\Greensoft\DrawioPpt`; `PACKAGE.txt` reports `v1.0.8 / Release / x64`.
- The Core, PowerPoint, and Word DLL SHA256 values in the standard directory match the release package. Both Office `CodeBase` values point to the standard install directory, with `LoadBehavior=3`.
- `verify-office-install.ps1`, `word-addin-load-check.ps1`, the installed-build `word-complex-metadata-e2e.ps1`, `word-selection-event-e2e.ps1`, and the actual Word URL-host regression passed independently. These are standard-install checks and are not counted as part of the temporary-installation 9/9 above.
- One 120-element complex SVG was used for a normal picture and a managed picture of the same size. The managed reference is 382 characters while the complete 254,606-character Draw.io XML is stored in `Document.CustomXMLParts`.
- Two crossed-order rounds ran in visible Word with a four-second warmup and 12-second measurement per picture per round; selection events and position updates were timed separately. Combined normal/managed P95 values were 7.501/5.833 ms for selection, 388.910/368.568 ms for position, and 112.280/103.247 ms for size. Managed/normal ratios were 0.778, 0.948, and 0.920. The comparison, absolute-latency, minimum-sample, final-failure, and responsiveness gates passed. The add-in was connected in the measured instance, the test counter classified 135 target-Shape events with zero missing, and installed-build runtime reflection reported `NoPassiveSelectionMetadataPath=True`.
- After save and reopen, the managed picture retained `69 / 82 / 340 / 226` for position and size, its lightweight reference remained 382 characters, the primary store restored 254,606 characters of XML, and `WordAddInConnect=True`.
- This sample supports only that no additional picture-metadata-path difference above the threshold was observed and that absolute latency remained inside this test's acceptance gates; it does not establish the cause of the absolute latency. See the [Word UI-thread stress acceptance report](./word-ui-thread-stress-report-v1.0.8.en.md) for the full command, method, and path-normalized evidence snapshot.

## Reports and Logs

- Summary report: `artifacts\test-reports\full-e2e-v1.0.8.md`
- Report copy: `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- PowerShell transcript: `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- Plug-in log snapshot: `artifacts\logs\v1.0.8\drawioppt-full-e2e.log`
- Standard-install UI-thread stress report: [v1.0.8 Word UI-thread stress acceptance report](./word-ui-thread-stress-report-v1.0.8.en.md)
- Machine-readable stress result: [path-normalized evidence snapshot](./evidence/word-ui-thread-stress-v1.0.8.json)
- Package contents and DLL hashes: [v1.0.8 release evidence](./release-evidence-v1.0.8.en.md)

## Acceptance Boundary

This report verifies the source Release DLLs, the temporary package installation, the standard local installation, and a real Word UI-thread stress comparison. The current Codex interactive desktop is denied cursor API access with Win32 error 5 and cannot perform a real continuous pointer drag. A manual 10-second pointer drag, resize, and reposition on each normal and managed picture, followed by a Re-edit button click, therefore remain a separate pending confirmation and are not counted in either the automated 9/9 or the UI-thread stress result.
