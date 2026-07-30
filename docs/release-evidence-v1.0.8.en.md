# v1.0.8 Release Evidence and Log Index

[English](./release-evidence-v1.0.8.en.md) | [中文](./release-evidence-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Deliverables

- Version: `v1.0.8`
- Release directory: `artifacts\releases\v1.0.8\package`
- Release archive: `artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip`
- Standard local install root: `%LOCALAPPDATA%\Greensoft\DrawioPpt`
- Version info file: `BuildInfo.txt`, with `BuildId` / `GitShortHash` also recorded in `PACKAGE.txt`
- New installed-package regressions: `install-release-upgrade-safety-test.ps1`, `word-svg-aspect-ratio-e2e.ps1`, `word-preview-image-provider-e2e.ps1`, and `powerpoint-svg-aspect-ratio-e2e.ps1`

v1.0.8 has been published as a GitHub Release: <https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.8>.

## Release Gates

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-sync-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\webview2-user-data-folder-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\version-display-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-release-upgrade-safety-test.ps1
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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" -DrawioSourcePath "<USER_DRAWIO_SOURCE>" -PictureWrapMode Front -PictureRenderFormat Png -PreviewPixelWidth 1240 -DurationSeconds 10 -MaximumP95Ratio 1.25 -MaximumPerRoundP95Ratio 1.25 -MaximumP95DeltaMs 50 -MaximumSelectionP95DeltaMs 50 -MaximumAbsoluteSelectionP95Ms 300 -MaximumAbsoluteP95Ms 300 -MaximumAbsoluteResizeP95Ms 300 -MinimumOperationsPerRound 10 -WarmupSeconds 4 -ResizeOperationsPerRound 12 -SelectionSettleMilliseconds 75 -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json"
```

## Final Verification Results

- The `Release|x64` build passed with 0 warnings and 0 errors.
- `version-display-safety-test.ps1` passed, covering `BuildInfo` lookup, Word/PPT Ribbon version labels, the shared settings-window version label, build/package generation of `BuildInfo.txt`, `BuildId` / `GitShortHash` in `PACKAGE.txt`, and installer output.
- `install-release-upgrade-safety-test.ps1` passed, covering the new package root equal to the old install root, the new package under a child folder of the old install root, and stale old files being removed after upgrade.
- Full functional Office E2E against the final package reported `13/13 PASS`: `ReleasePackage`, `NestedInstallUpgrade`, `InstallRelease`, `InstalledOfficeAddInsVerify`, `PowerPointAddInLoad`, `InstalledUrlSmoke`, `InstalledWordUrlHostE2E`, `InstalledWordComplexMetadataE2E`, `InstalledWordSvgAspectE2E`, `InstalledWordPreviewProviderE2E`, `InstalledPowerPointSvgAspectE2E`, `PowerPointUrlE2E`, and `DesktopExporter` all passed.
- The full-E2E summary and log copy are content-identical; the final rerun regenerates `artifacts\test-reports\full-e2e-v1.0.8.md` and `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`.
- The pre-compression internal Markdown-link gate reported `PackageMarkdownMissingLinkCount=0`.
- Full E2E identifies only test-owned Office processes by PID, start time, and process name; it does not stop pre-existing user processes. It snapshots Word and PowerPoint registration trees before installation, then restores values, types, subkeys, and originally absent state in `finally`. Residual `WINWORD` / `POWERPNT` PIDs enter FatalError and fail E2E.
- `NestedInstallUpgrade` uses the real release package to validate the upgrade path where the new package is inside the old install root, confirming that the installer does not delete its own source and does not keep stale files from the previous installation.
- Temporary-package uninstall, registration-state restoration, and settings restoration succeeded. The run reported `FullE2ECleanupSucceeded=True`, with no residual Office process.
- Latest-source Word normally displays a 1240px aspect-preserving PNG as a floating `wdWrapFront` picture, with horizontal and vertical diagrams using `contain`. The enhanced real regression reports `VerticalRatio=0.25`, `VerticalContained=True`, and `FailedReplacementPreservedOriginal=True`, proving create-new-before-delete-old replacement. Export failure generates a lightweight aspect-preserving PNG placeholder from a 1240px baseline, caps extreme portrait height at 1200px, and does not use complex SVG; temporary XML implementation includes retry, deferred, and startup cleanup. Only an explicit command after selection reads the live selection and starts editing.
- PowerPoint inserts the original SVG directly, centers it with source-proportional `contain` sizing inside the slide bounds, and neither rewrites the `viewBox` nor introduces internal blank padding.

## DLL SHA256 and ZIP Recording Boundary

The following hashes were re-read on 2026-07-29 after final packaging and standard local installation. All three DLLs match across source Release output, the release package, and the standard local installation.

| File | Source Release | Release package | Standard local installation |
| --- | --- | --- | --- |
| `DrawioPpt.Core.dll` | `AF5025768052D450D7B43E70229CA4FAEFECEBA50C3D1C33157F4FA7C560AED7` | `AF5025768052D450D7B43E70229CA4FAEFECEBA50C3D1C33157F4FA7C560AED7` | `AF5025768052D450D7B43E70229CA4FAEFECEBA50C3D1C33157F4FA7C560AED7` |
| `DrawioPpt.PowerPointAddIn.dll` | `4717622B7F8995F437989E3D9C2FF14A33C918EA9848A4702ADAFA095753931F` | `4717622B7F8995F437989E3D9C2FF14A33C918EA9848A4702ADAFA095753931F` | `4717622B7F8995F437989E3D9C2FF14A33C918EA9848A4702ADAFA095753931F` |
| `DrawioPpt.WordAddIn.dll` | `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8` | `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8` | `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8` |

Relative locations:

- Source Release: `src\<project>\bin\x64\Release\<dll>`
- Release package: `artifacts\releases\v1.0.8\package\bin\<dll>`
- Standard local installation: `%LOCALAPPDATA%\Greensoft\DrawioPpt\bin\<dll>`

The final package contains Word DLL `A6AB…492E8` and has been installed over the standard local installation. Installed Word, preview-provider, and PowerPoint aspect-ratio E2E tests passed again.

ZIP SHA256 is not embedded in this evidence document because the document is packaged into that same ZIP and would create a self-reference. The release script or an external delivery summary records the hash after final packaging.

## Word UI-thread Stress Evidence

- First corrected result: `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json`, `ReportVersion=9`, SHA256 `0C4B1FFB8C2ABB31505B277A1035E2716D87E7456BF47C7030DBFD91618A5B1B`.
- Confirmation rerun: `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-rerun.json`, `ReportVersion=9`, SHA256 `0D153F545245F0F0D6E405AA717E22012FE6916724722D14B6E111B6254F7D16`.
- Latest result after opaque-raster detection was corrected: `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-final.json`, `ReportVersion=9`, SHA256 `4C7B7C5EDA65B7F5BBF0A56FE18D974C7AA863B1C84D7E3556DF58D355A90C42`.
- Test picture: a 1240×1753 PNG. The independent SVG source ratio is `0.707123`, the PNG ratio is `0.707359`, the error is `0.000236`, and layout is floating `Front`. The PNG is opaque 24bpp, so transparent-padding inspection explicitly reports `BlankPaddingDetectionSupported=false` and null `BorderPaddingPassed` instead of claiming a pixel-level pass; export uses `--border 0`. Normal and managed pictures have matching dimensions, anchor semantics, and wrapping.
- Latest-run normal/managed selection P95: `68.866/74.243 ms`; position P95: `321.844/325.764 ms`; resize P95: `89.633/122.968 ms`; target selection events: `96/96`, with zero missing.
- The latest run reports `ComparisonPassed=false` and `NoAdditionalMetadataPathDifferenceObserved=false`. Managed-picture round position P95 values were `303.892/337.339 ms`, while the normal control reported `328.256/318.748 ms`. An earlier confirmation run passed the relative gates, but it cannot override the latest failure.
- `AbsoluteLatencyGatePassed=false` and `OverallPassed=false`; the latest run had zero final failures and zero Word-unresponsive samples.
- `PointerInputCovered=false`. Automation does not inject real mouse input. This evidence makes no successful mouse-drag claim; the user chose to skip that manual acceptance, so its status is not “passed.”
- The release-facing [sanitized machine-readable summary](./evidence/word-ui-thread-stress-v1.0.8.json) removes user names, absolute attachment paths, temporary GUIDs, and process identities.

## Reports and Logs

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-rerun.json`
- `artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px-final.json`
- [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md)
- [Word UI-thread stress test report](./word-ui-thread-stress-report-v1.0.8.en.md)
- [Sanitized Word UI-thread stress evidence](./evidence/word-ui-thread-stress-v1.0.8.json)

## Known Boundary

The final package, standard local installation, and latest source DLLs are synchronized, and functional `13/13 PASS` covers this Word refinement plus the installed-root upgrade fix. The latest independent stress run has managed and normal-control rounds above 300ms and also fails the relative gate, so `OverallPassed=false`; the user chose to skip real-pointer acceptance, and it must not be represented as passed. Record the final ZIP hash outside the archive.
