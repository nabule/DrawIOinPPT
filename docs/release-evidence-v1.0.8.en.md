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
- Full functional Office E2E against the final package reported `12/12 PASS`: `ReleasePackage`, `InstallRelease`, `InstalledOfficeAddInsVerify`, `PowerPointAddInLoad`, `InstalledUrlSmoke`, `InstalledWordUrlHostE2E`, `InstalledWordComplexMetadataE2E`, `InstalledWordSvgAspectE2E`, `InstalledWordPreviewProviderE2E`, `InstalledPowerPointSvgAspectE2E`, `PowerPointUrlE2E`, and `DesktopExporter` all passed.
- The full-E2E summary and log copy are content-identical, with SHA256 `7D5AC38245C29AF08B04C24414480E1024E270210657EC0AFB97D0ABF52CF73C`. The transcript SHA256 is `3006C8C50E311F8F5BB65C9745BE131A440215FFB3B5A0780F96B1D773A37121`.
- The pre-compression internal Markdown-link gate reported `PackageMarkdownMissingLinkCount=0`.
- Full E2E identifies only test-owned Office processes by PID, start time, and process name; it does not stop pre-existing user processes. It snapshots Word and PowerPoint registration trees before installation, then restores values, types, subkeys, and originally absent state in `finally`. Residual `WINWORD` / `POWERPNT` PIDs enter FatalError and fail E2E.
- Temporary-package uninstall, registration-state restoration, and settings restoration succeeded. The run reported `FullE2ECleanupSucceeded=True`, with no residual Office process.
- Latest-source Word normally displays a 620px aspect-preserving PNG as a floating `wdWrapFront` picture, with horizontal and vertical diagrams using `contain`. The enhanced real regression reports `VerticalRatio=0.25`, `VerticalContained=True`, and `FailedReplacementPreservedOriginal=True`, proving create-new-before-delete-old replacement. Export failure generates a lightweight aspect-preserving 620×310 PNG placeholder instead of complex SVG; temporary XML uses retry, deferred, and startup cleanup. Only an explicit command after selection reads the live selection and starts editing.
- PowerPoint inserts the original SVG directly, centers it with source-proportional `contain` sizing inside the slide bounds, and neither rewrites the `viewBox` nor introduces internal blank padding.

## DLL SHA256 and ZIP Recording Boundary

The following hashes were re-read on 2026-07-29 after final packaging and standard local installation. All three DLLs match across source Release output, the release package, and the standard local installation.

| File | Source Release | Release package | Standard local installation |
| --- | --- | --- | --- |
| `DrawioPpt.Core.dll` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` |
| `DrawioPpt.PowerPointAddIn.dll` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` | `EB500C2575C929BDD6CAA7718A47A7671D604C468C77CF197497CEE1096F1F3E` |
| `DrawioPpt.WordAddIn.dll` | `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D` | `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D` | `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D` |

Relative locations:

- Source Release: `src\<project>\bin\x64\Release\<dll>`
- Release package: `artifacts\releases\v1.0.8\package\bin\<dll>`
- Standard local installation: `%LOCALAPPDATA%\Greensoft\DrawioPpt\bin\<dll>`

The final package contains Word DLL `6E6D…127D` and has been installed over the standard local installation. Installed Word, preview-provider, and PowerPoint aspect-ratio E2E tests passed again.

ZIP SHA256 is not embedded in this evidence document because the document is packaged into that same ZIP and would create a self-reference. The release script or an external delivery summary records the hash after final packaging.

## Word UI-thread Stress Evidence

- Raw result: `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json`, `ReportVersion=9`, SHA256 `935D7C0827E872B5F29309EC1A30E0FB72898865CB4092BDBDAFE6B35928DCB1`.
- Test picture: a 620×876 PNG that preserves source ratio with no border, using floating `Front` layout. Normal and managed pictures have matching dimensions, anchor semantics, and wrapping.
- Normal/managed position P95: `524.962/488.881 ms`; managed-minus-normal delta: `-36.081 ms`; target selection events: `98/98`, with zero missing.
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

The final package, standard local installation, and latest source DLLs are synchronized, and functional `12/12 PASS` covers this Word refinement. The separate stress run remains `OverallPassed=false` because of the strict absolute `300 ms` gate and missing real-pointer coverage. The user must still perform real-pointer drag, resize, reposition, and a manual Re-edit click. Record the final ZIP hash outside the archive.
