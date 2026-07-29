# v1.0.8 Release Evidence and Log Index

[English](./release-evidence-v1.0.8.en.md) | [中文](./release-evidence-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Deliverables

- Version: `v1.0.8`
- Release directory: `artifacts\releases\v1.0.8\package`
- Release archive: `artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip`
- Standard local install root: `%LOCALAPPDATA%\Greensoft\DrawioPpt`
- New installed-package regressions: `word-svg-aspect-ratio-e2e.ps1`, `word-preview-image-provider-e2e.ps1`, and `powerpoint-svg-aspect-ratio-e2e.ps1`

v1.0.8 remains pending public release; this evidence does not mean that a GitHub Release has been created.

## Release Gates

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-sync-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\webview2-user-data-folder-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-cleanup-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-preview-image-provider-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-svg-aspect-ratio-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\powerpoint-svg-aspect-ratio-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" -DrawioSourcePath "<USER_DRAWIO_SOURCE>" -PictureWrapMode Front -PictureRenderFormat Png -PreviewPixelWidth 620 -DurationSeconds 10 -MaximumP95Ratio 1.25 -MaximumPerRoundP95Ratio 1.25 -MaximumP95DeltaMs 50 -MaximumSelectionP95DeltaMs 50 -MaximumAbsoluteSelectionP95Ms 300 -MaximumAbsoluteP95Ms 300 -MaximumAbsoluteResizeP95Ms 300 -MinimumOperationsPerRound 10 -WarmupSeconds 4 -ResizeOperationsPerRound 12 -SelectionSettleMilliseconds 75 -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json"
```

## Final Verification Results

- The `Release|x64` build passed with 0 warnings and 0 errors.
- Full functional Office E2E against the existing temporary package snapshot reported `12/12 PASS`: `ReleasePackage`, `InstallRelease`, `InstalledOfficeAddInsVerify`, `PowerPointAddInLoad`, `InstalledUrlSmoke`, `InstalledWordUrlHostE2E`, `InstalledWordComplexMetadataE2E`, `InstalledWordSvgAspectE2E`, `InstalledWordPreviewProviderE2E`, `InstalledPowerPointSvgAspectE2E`, `PowerPointUrlE2E`, and `DesktopExporter` all passed.
- The full-E2E summary and log copy are content-identical, with SHA256 `BA3A00156E3ED56CC15C57EE175FD8E7A524D0B7662649E8A093CFCD81D4513F`. The transcript SHA256 is `F1FD957A318FD23809CD20E7D9722F11CCE34084169F1FD49E21BCCBE9489A7E`.
- The pre-compression internal Markdown-link gate reported `PackageMarkdownMissingLinkCount=0`.
- Full E2E identifies only test-owned Office processes by PID, start time, and process name; it does not stop pre-existing user processes. It snapshots Word and PowerPoint registration trees before installation, then restores values, types, subkeys, and originally absent state in `finally`. Residual `WINWORD` / `POWERPNT` PIDs enter FatalError and fail E2E.
- Temporary-package uninstall, registration-state restoration, and settings restoration succeeded. The run reported `FullE2ECleanupSucceeded=True`, with no residual Office process.
- Latest-source Word normally displays a 620px aspect-preserving PNG as a floating `wdWrapFront` picture, with horizontal and vertical diagrams using `contain`. The enhanced real regression reports `VerticalRatio=0.25`, `VerticalContained=True`, and `FailedReplacementPreservedOriginal=True`, proving create-new-before-delete-old replacement. Export failure generates a lightweight aspect-preserving 620×310 PNG placeholder instead of complex SVG; temporary XML uses retry, deferred, and startup cleanup. Only an explicit command after selection reads the live selection and starts editing.
- PowerPoint inserts the original SVG directly, centers it with source-proportional `contain` sizing inside the slide bounds, and neither rewrites the `viewBox` nor introduces internal blank padding.

## DLL SHA256 and ZIP Recording Boundary

The following hashes were re-read on 2026-07-29 after the latest source regression. Core and PowerPoint match across all three locations; the source Word DLL has changed, while the package and standard local installation have not yet been synchronized.

| File | Source Release | Release package | Standard local installation |
| --- | --- | --- | --- |
| `DrawioPpt.Core.dll` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` |
| `DrawioPpt.PowerPointAddIn.dll` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` |
| `DrawioPpt.WordAddIn.dll` | `F5B68A12E0E7CBB54EEBE501239F45928EB4DFB685E8007AA57ABBE93615AD70` | `A888A38D2B45DC2126510A2B63669DA89812E4C7580E03EAC5DD02285045B5D9` | `A888A38D2B45DC2126510A2B63669DA89812E4C7580E03EAC5DD02285045B5D9` |

Relative locations:

- Source Release: `src\<project>\bin\x64\Release\<dll>`
- Release package: `artifacts\releases\v1.0.8\package\bin\<dll>`
- Standard local installation: `%LOCALAPPDATA%\Greensoft\DrawioPpt\bin\<dll>`

The current ZIP still carries the pre-document-closure v8 stress-evidence copy and the `A888…B5D9` Word DLL; it does not contain the latest `F5B6…AD70` source build and must not be used as a current Word-DLL stress-pass record. Rebuild the package, update the standard local installation, rerun full E2E, and revalidate package contents and links.

ZIP SHA256 is not embedded in this evidence document because the document is packaged into that same ZIP and would create a self-reference. The release script or an external delivery summary records the hash after final packaging.

## Word UI-thread Stress Evidence

- Raw result: `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json`, `ReportVersion=9`, SHA256 `82CBC3A18C43495B1C24154D6FE28981B7E282633C073592816AABC4BEB2EC54`.
- Test picture: a 620×876 PNG that preserves source ratio with no border, using floating `Front` layout. Normal and managed pictures have matching dimensions, anchor semantics, and wrapping.
- Normal/managed position P95: `556.177/574.474 ms`; managed-minus-normal delta: `18.297 ms`; target selection events: `94/94`, with zero missing.
- The absolute position-P95 gate of `300 ms` failed. A 17×24 solid-color PNG baseline also failed it.
- `SelectionComparisonPassed=false`, `ComparisonPassed=false`, `AbsoluteLatencyGatePassed=false`, `NoAdditionalMetadataPathDifferenceObserved=false`, and `OverallPassed=false`.
- `PointerInputCovered=false`. Windows denied `GetCursorPos`, and `SetIsBorderRequired` is unsupported. This evidence makes no successful mouse-drag claim.
- The release-facing [sanitized machine-readable summary](./evidence/word-ui-thread-stress-v1.0.8.json) removes user names, absolute attachment paths, temporary GUIDs, and process identities.

## Reports and Logs

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json`
- [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md)
- [Word UI-thread stress test report](./word-ui-thread-stress-report-v1.0.8.en.md)
- [Sanitized Word UI-thread stress evidence](./evidence/word-ui-thread-stress-v1.0.8.json)

## Known Boundary

The existing package snapshot's `12/12 PASS` does not cover the later Word source refinements, real mouse input, or the separate stress-gate failures. Before public release, rebuild/reinstall and rerun full E2E; the user must still perform real-pointer drag, resize, reposition, and a manual Re-edit click. Record the final ZIP hash outside the archive.
