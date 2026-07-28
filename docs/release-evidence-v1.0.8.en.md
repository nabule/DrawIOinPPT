# v1.0.8 Release Evidence and Log Index

[English](./release-evidence-v1.0.8.en.md) | [中文](./release-evidence-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Deliverables

- Version: `v1.0.8`
- Release directory: `artifacts\releases\v1.0.8\package`
- Release archive: `artifacts\releases\v1.0.8\DrawioPpt-v1.0.8.zip`
- New installed-package regression: `scripts\word-complex-metadata-e2e.ps1`

v1.0.8 remains pending public release; this repository evidence does not mean that a GitHub Release has been created.

## Release Gates

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-sync-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\webview2-user-data-folder-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.8 -SkipBuild
```

## Final Verification Results

- The `Release|x64` build passed with 0 warnings and 0 errors.
- All source-contract and real Word pre-release gates passed; see the [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md) for the detailed evidence.
- The full Office E2E against the temporary release-package installation passed 9/9: `ReleasePackage`, `InstallRelease`, `InstalledOfficeAddInsVerify`, `PowerPointAddInLoad`, `InstalledUrlSmoke`, `InstalledWordUrlHostE2E`, `InstalledWordComplexMetadataE2E`, `PowerPointUrlE2E`, and `DesktopExporter` all reported PASS.
- The package contains `DrawioPpt.WordAddIn.dll`, `scripts\word-complex-metadata-e2e.ps1`, and the versioned Chinese and English release notes, E2E reports, and evidence indexes.
- The full E2E restored the repository Release add-in registration and left no `WINWORD` or `POWERPNT` process behind.

## Release-package DLL SHA256

| File | SHA256 |
| --- | --- |
| `bin\DrawioPpt.Core.dll` | `C851D64AD79689EF83615E619C6A2CFD3E07599AD2EC888602A3F4136D8FADA6` |
| `bin\DrawioPpt.PowerPointAddIn.dll` | `C8425FC93C103C3CFAA80A0BF33468363ADE95165EE76308B98A5FA37F76786A` |
| `bin\DrawioPpt.WordAddIn.dll` | `4FE2ECE68A8F3762A5B930AB814DF64A41766D11FCA33DA12F8BBF3DF51A5BB8` |

The archive SHA256 is not embedded in a document that is packaged into the same archive, avoiding a self-hash loop. It is recorded separately after the final package rebuild.

## Reports and Logs

- `artifacts\test-reports\full-e2e-v1.0.8.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.report.md`
- `artifacts\logs\v1.0.8\full-e2e-v1.0.8.transcript.log`
- `artifacts\logs\v1.0.8\drawioppt-full-e2e.log`

## Known Preconditions and Boundary

Close Word and PowerPoint before running real Office tests, and do not run another Office automation task in parallel. Final installation into the standard local directory and the manual drag comparison remain later acceptance work and are not claimed here.
