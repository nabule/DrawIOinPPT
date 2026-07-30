# v1.0.8 E2E Test Report

[English](./e2e-test-report-v1.0.8.en.md) | [中文](./e2e-test-report-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Test Environment

- Date: 2026-07-29
- Windows + Microsoft Word Desktop / PowerPoint Desktop
- .NET Framework 4.8, `Release|x64`
- Current-user Office COM Add-in registration, an isolated temporary v1.0.8 package installation, and the standard local installation

## Pre-release Gates

Real Office commands ran serially. No user Word or PowerPoint process could be present before they started, and no `WINWORD` or `POWERPNT` process could remain afterward.

| Check | Result | Evidence |
| --- | --- | --- |
| Release x64 build | PASS | 0 warnings, 0 errors |
| Word zero-passive-selection-path source contract | PASS | `WORD_SELECTION_SYNC_TEST_PASS`: no `WindowSelectionChange`, auto-open, or startup selection read; each of the four explicit commands captures one live picture |
| WebView2 user-data-folder source contract | PASS | `WEBVIEW2_USER_DATA_FOLDER_SOURCE_TEST_PASS` |
| Installer upgrade safety contract | PASS | `INSTALL_RELEASE_UPGRADE_SAFETY_TEST_PASS`: when the new package is run from the old install root or a child folder, it does not delete its own source, and stale old files are removed |
| Word URL-host safety contract | PASS | `WORD_URL_ADDIN_HOST_E2E_SAFETY_TEST_PASS` |
| Full-E2E cleanup safety contract | PASS | `FULL_E2E_CLEANUP_SAFETY_TEST_PASS`: snapshots Word/PowerPoint registration trees before installation; `finally` restores values, types, subkeys, and originally absent state; residual Office PIDs enter FatalError and fail E2E |
| Word complex-metadata E2E | PASS | Document-level primary storage, lightweight reference, failure fallback, legacy migration, stable part ID, and save-close-reopen persistence satisfied their assertions |
| Word explicit-selection recognition E2E | PASS | `NoPassiveSelectionMetadataPath=True`, `ExplicitManagedSelectionDetected=True`, `ExplicitPlainPictureCanBind=True` |
| Real Word URL-host E2E | PASS | `WordAddInConnect=True`, `ExplicitEditAutomationAvailable=True`, `ExplicitEditCommandInvoked=True`, `ActualWordUrlEditorSaved=True` |

Explicit-command automation verifies select-then-Re-edit. It is neither selection-triggered auto-open nor acceptance of a real user mouse click.

## Full Office E2E Against the Temporary Installation

Command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

Result: `13/13 PASS`.

| Check | Result | Evidence |
| --- | --- | --- |
| `ReleasePackage` | PASS | The v1.0.8 release directory was created |
| `NestedInstallUpgrade` | PASS | The real release package was run from a child folder under the old install root; the installer did not delete its own source and stale old files were removed |
| `InstallRelease` | PASS | The package was installed into an isolated temporary directory |
| `InstalledOfficeAddInsVerify` | PASS | Word / PowerPoint registration and COM loading passed |
| `PowerPointAddInLoad` | PASS | `Connect=True` |
| `InstalledUrlSmoke` | PASS | URL-editor smoke completed save and SVG/XML write-back |
| `InstalledWordUrlHostE2E` | PASS | Temporary-package DLLs completed URL write-back through select-then-explicit-command in real `WINWORD.EXE` |
| `InstalledWordComplexMetadataE2E` | PASS | Complex-XML primary storage, lightweight reference, fallback, and migration passed |
| `InstalledWordSvgAspectE2E` | PASS | The package snapshot passed Word 2:1 create, replacement, and legacy-ratio correction with floating `wdWrapFront` |
| `InstalledWordPreviewProviderE2E` | PASS | The package snapshot passed Word 1240px aspect-preserving PNG, no forced square, temporary-source cleanup, and safe fallback |
| `InstalledPowerPointSvgAspectE2E` | PASS | PowerPoint original-SVG `contain` sizing and no `viewBox` expansion passed |
| `PowerPointUrlE2E` | PASS | PowerPoint URL create, reopen, edit, and persistence passed |
| `DesktopExporter` | PASS | draw.io Desktop exported SVG successfully |

Packaging reported `PackageMarkdownMissingLinkCount=0`. The summary and log copy are content-identical; final rerun hashes are recorded in the [release evidence](./release-evidence-v1.0.8.en.md).

After the temporary test installation was uninstalled, cleanup restored Word and PowerPoint add-in registration to the exact pre-run state, including keys that were absent before the run. Settings restoration completed, `FullE2ECleanupSucceeded=True`, and no `WINWORD` or `POWERPNT` process remained.

The final package includes the enhanced real-Word regressions: `VerticalRatio=0.25`, `VerticalContained=True`, and `FailedReplacementPreservedOriginal=True`. The preview gate covers the normal 1240px PNG, a 1240×620 aspect-preserving lightweight PNG placeholder for its 2:1 failure sample, and immediate temporary-source, preview, and fallback-directory cleanup. The implementation still includes bounded retry, deferred background cleanup, and startup cleanup for temporary XML, but this gate does not fault-inject those three paths. Complex-SVG fallback is no longer used.

## Standard Local Installation and Artifact Consistency

- Standard install root: `%LOCALAPPDATA%\Greensoft\DrawioPpt`; `PACKAGE.txt` reports `v1.0.8 / Release / x64`.
- Core, PowerPoint, and Word DLL SHA256 values match across source Release output, the release package, and the standard local installation. The Word DLL is `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8`.
- PowerPoint inserts the original SVG directly, centers it with proportional `contain` sizing inside the slide bounds, and neither rewrites the `viewBox` nor introduces internal blank padding.
- In the latest source, Word normally displays a 1240px aspect-preserving PNG as a floating `wdWrapFront` `Shape`; horizontal and vertical diagrams use `contain`, replacement creates the new picture before deleting the old one, and failures preserve the original. Export failure creates a lightweight aspect-preserving PNG placeholder from a 1240px baseline and caps extreme portrait height at 1200px while source XML remains in document metadata. Selection alone does not edit; only an explicit command reads the selection and starts editing.
- ZIP SHA256 is not embedded in this report because the report is packaged into that same archive. The release script or an external delivery summary records it after final repackaging.

## Word UI-thread Stress Result

The latest user-complex-source run after blank-padding validation was corrected used a 1240×1753 PNG, floating `Front` layout, and a rich-document baseline. Normal/managed selection P95 values were `68.866/74.243 ms`, position P95 values were `321.844/325.764 ms`, and resize P95 values were `89.633/122.968 ms`; all `96/96` target-Shape selection events were confirmed, with none missing. Actual pixel width, independent source ratio, save/reopen, geometry, storage, and cleanup gates passed.

The latest run reports `ComparisonPassed=false`, `NoAdditionalMetadataPathDifferenceObserved=false`, `AbsoluteLatencyGatePassed=false`, and `OverallPassed=false`. Managed-picture round position P95 values were `303.892/337.339 ms`, while the normal control reported `328.256/318.748 ms`. Every round had zero final set failures and zero Word-unresponsive samples, but this is still not a complete automated performance-acceptance pass.

The latest PNG is opaque 24bpp, so transparent-padding pixel inspection is not applicable and explicitly reports `BlankPaddingDetectionSupported=false`; it is not represented as a pixel-level pass. Aspect preservation is evidenced by the independent SVG source ratio and a PNG-ratio error of `0.000236`, while export explicitly uses `--border 0`. Focused counterexamples separately verify detection of transparent padding and preservation of a 1px edge-touching line.

`PointerInputCovered=false`. Automation updates `Shape` properties on Word's visible STA UI thread and does not inject real mouse input, so no successful mouse-drag claim is made. The user chose to skip that manual acceptance; its status is “skipped/not accepted,” not “passed.”

## Reports and Logs

- Functional E2E summary: `artifacts\test-reports\full-e2e-v1.0.8.md`
- Log copy: `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- PowerShell transcript: `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- Raw 1240px Word stress result: `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json`
- Latest result after corrected blank-padding validation: `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-final.json`
- [Word UI-thread stress test report](./word-ui-thread-stress-report-v1.0.8.en.md)
- [Sanitized machine-readable stress evidence](./evidence/word-ui-thread-stress-v1.0.8.json)
- [Release contents and hashes](./release-evidence-v1.0.8.en.md)

## Acceptance Boundary

This report records `13/13 PASS` for the final package, and the standard local installation was updated and reverified; the added gate covers upgrading from inside an existing install root. Word stress gates are separate: managed-picture rounds stayed below 300ms, but a normal-control round exceeded the gate, so `OverallPassed=false`. The user chose to skip real-pointer drag, resize, reposition, and a manual Re-edit click; skipped must not be represented as passed.
