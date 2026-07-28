# v1.0.8 Release Evidence and Log Index

[English](./release-evidence-v1.0.8.en.md) | [中文](./release-evidence-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Deliverables

- Version: `v1.0.8`
- Release directory: `artifacts\releases\v1.0.8\package`
- Release archive: `artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip`
- Standard local install root: `%LOCALAPPDATA%\Greensoft\DrawioPpt`
- New installed-package regression: `scripts\word-complex-metadata-e2e.ps1`

v1.0.8 remains pending public release; this repository evidence does not mean that a GitHub Release has been created.

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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" -DurationSeconds 12 -MaximumP95Ratio 1.25 -MaximumPerRoundP95Ratio 1.50 -MaximumSelectionP95DeltaMs 100 -MaximumAbsoluteSelectionP95Ms 2000 -MaximumAbsoluteP95Ms 2000 -MaximumAbsoluteResizeP95Ms 750 -MinimumOperationsPerRound 10 -WarmupSeconds 4 -ResizeOperationsPerRound 12 -SelectionSettleMilliseconds 75 -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-v1.0.8.json"
```

## Final Verification Results

- The `Release|x64` build passed with 0 warnings and 0 errors.
- All source-contract and real Word pre-release gates passed; see the [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md) for the detailed evidence.
- The full Office E2E against the temporary release-package installation passed 9/9: `ReleasePackage`, `InstallRelease`, `InstalledOfficeAddInsVerify`, `PowerPointAddInLoad`, `InstalledUrlSmoke`, `InstalledWordUrlHostE2E`, `InstalledWordComplexMetadataE2E`, `PowerPointUrlE2E`, and `DesktopExporter` all reported PASS.
- The package contains `DrawioPpt.WordAddIn.dll`, `scripts\word-complex-metadata-e2e.ps1`, both architecture documents, both regression checklists, both optimization backlogs, the historical v1.0.5 E2E reports, and the versioned release notes, E2E reports, and evidence indexes. `PACKAGE.txt` also lists both README files and itself.
- The pre-compression package-link gate reported `PackageMarkdownMissingLinkCount=0`.
- Full E2E identifies its test-owned Word process by PID and start time. The main script AST has only one trailing `exit 1`; build, package, install, and compilation failures throw into the common catch/finally path. A controlled probe verifies that the main error plus uninstall, repository re-registration, and settings-restoration errors are all preserved in the FatalError row before the unified exit returns 1.
- In the final full E2E run, temporary-package uninstall, repository Release add-in registration restoration, and settings restoration all succeeded. The run reported `FullE2ECleanupSucceeded=True` and left no `WINWORD` or `POWERPNT` process behind.
- The repository-registration restoration in the previous item is the historical end state of the temporary-installation E2E. The later standard local installation switched both Word and PowerPoint `CodeBase` values to `%LOCALAPPDATA%\Greensoft\DrawioPpt\bin`, kept `LoadBehavior=3`, and passed real COM loading checks.
- All three DLL SHA256 values in the standard installation match the release package. Installed-build complex-metadata, zero-passive-selection-path and explicit-recognition, and actual Word URL-host regressions passed. The real host reported `ExplicitEditAutomationAvailable=True` and `ExplicitEditCommandInvoked=True`, confirming select-then-explicit-Re-edit rather than selection-triggered auto-open.
- The installed Word UI-thread stress comparison ran two crossed-order rounds with a four-second warmup and 12-second measurement per picture per round, recording selection, position, and size latency separately. Combined normal/managed P95 values were 7.501/5.833 ms, 388.910/368.568 ms, and 112.280/103.247 ms, with managed/normal ratios of 0.778, 0.948, and 0.920. The comparison, absolute-latency, minimum-sample, final-failure, and responsiveness gates passed. The add-in was connected in the measured instance, 135 target-Shape events were classified with zero missing, and installed-build runtime reflection reported `NoPassiveSelectionMetadataPath=True`. Save and reopen restored the 382-character lightweight reference and all 254,606 source-XML characters. This sample does not establish the cause of the absolute latency.

## Release-package DLL SHA256

| File | SHA256 |
| --- | --- |
| `bin\DrawioPpt.Core.dll` | `5A7E8E57B485D39FCB4B46D23A3DBB0ACDCEBB212B48A0A39543E3C9F104C11E` |
| `bin\DrawioPpt.PowerPointAddIn.dll` | `09E4BBE61B905BACA7B924868696A724A59EF37D3417F58E580AB7A01894733A` |
| `bin\DrawioPpt.WordAddIn.dll` | `6587C12F23709E5102DE5DD13DBD647CDA45CBDFB778E5D82FA12FA70761219D` |

The archive SHA256 is not embedded in a document that is packaged into the same archive, avoiding a self-hash loop. It is recorded separately after the final package rebuild.

## Reports and Logs

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\logs\v1.0.8\drawioppt-full-e2e.log`
- [v1.0.8 Word UI-thread stress acceptance report](./word-ui-thread-stress-report-v1.0.8.en.md)
- [Word UI-thread stress path-normalized evidence snapshot](./evidence/word-ui-thread-stress-v1.0.8.json)

## Known Preconditions and Boundary

Close Word and PowerPoint before running real Office tests, and do not run another Office automation task in parallel. Standard local installation and the real Word UI-thread stress comparison are complete. The current interactive desktop cannot inject pointer input because the operating system returns Win32 error 5, so a manual 10-second pointer drag, resize, and reposition on each normal and managed picture, followed by a Re-edit button click, remain for user confirmation. This evidence does not describe the automated stress result as manual mouse acceptance.
