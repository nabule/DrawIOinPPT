# Full E2E Test Report (v1.0.0)

[English](./e2e-test-report-v1.0.0.en.md) | [中文](./e2e-test-report-v1.0.0.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This report records the real release validation results for `DrawioPpt v1.0.0`. The test execution date was `2026-04-12`. The goal was to confirm that the release package, installation scripts, PowerPoint host loading, URL-mode write-back flow, editing again after reopening, and desktop export capability all work on the current machine.

## Test Environment

- OS: `Microsoft Windows 10 Pro 10.0.19045 64-bit`
- PowerPoint: `Microsoft PowerPoint Desktop x64`
- .NET: `.NET Framework 4.8`
- WebView2 Runtime: installed
- draw.io Desktop: `C:\Program Files\draw.io\draw.io.exe`
- Release version: `v1.0.0`
- Validated commit: `6239eab246aa603a90af119b640aaff1186f95dc`

## Commands Actually Executed

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1 -SkipBuild
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1 -Version v1.0.0 -SkipBuild
```

## Coverage

- Environment prerequisite check
- Release build
- Local smoke test for the URL editor
- Release package generation
- Release package installation and current-user registration
- Automatic loading of the PowerPoint COM add-in
- URL-mode new diagram flow including `configure -> init -> load -> save -> export`
- Saving the PPT, closing it, reopening it, and editing the same diagram again
- Persistence checks for `Shape.Tags / AlternativeText / Presentation.CustomXMLParts`
- draw.io Desktop CLI SVG export verification

## Key Observations

- URL smoke test returned:
  - `DialogResult=OK`
  - `Saved=True`
  - `SvgHasSmoke=True`
  - `XmlHasSmokeId=True`
- After the initial creation during the full E2E run:
  - `CreatedManagedShape=True`
  - `CreatedDiagramId=7a0c3b5b674d4934a99a74c6f7546543`
  - `CreatedCustomXmlCount=4`
- After closing, reopening, and editing again:
  - `ReopenedEditApplied=True`
  - `EditedCustomXmlCount=4`
- After closing and reopening again for persistence validation:
  - `PersistedAfterReopen=True`
  - `FinalCustomXmlCount=4`
  - `ShapeCount=1`

These results show that, on the current machine, the `v1.0.0` URL-mode write-back flow, second-round editing, and `CustomXMLParts` persistence path have all passed validation on a real PowerPoint host.

## Summary Table

| Check | Result | Detail |
| --- | --- | --- |
| ReleasePackage | PASS | `artifacts\releases\v1.0.0\package` |
| InstallRelease | PASS | The release package installs into a temporary install directory and completes registration |
| PowerPointAddInLoad | PASS | `Connect=True` |
| InstalledUrlSmoke | PASS | `url-editor-smoke.ps1` in the release package runs successfully |
| PowerPointUrlE2E | PASS | `ExitCode=0` |
| DesktopExporter | PASS | `C:\Program Files\draw.io\draw.io.exe` |

## Artifacts and Logs

- Raw E2E result: `artifacts\test-reports\full-e2e-v1.0.0.md`
- Log directory for this version: `artifacts\logs\v1.0.0\`
- Add-in log snapshot: `artifacts\logs\v1.0.0\drawioppt-full-e2e.log`
- E2E console log: `artifacts\logs\v1.0.0\full-e2e.console.log`
- URL smoke-test log: `artifacts\logs\v1.0.0\url-editor-smoke.log`

## Conclusion

`v1.0.0` meets the release criteria in the current validation environment: the program builds successfully, the release package can be produced and installed, PowerPoint loads the add-in correctly, the URL-mode create/write-back/reopen/re-edit flow passes real testing, and desktop export capability has also been verified.
