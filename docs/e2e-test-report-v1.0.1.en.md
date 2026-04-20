# Full E2E Test Report (v1.0.1)

[English](./e2e-test-report-v1.0.1.en.md) | [中文](./e2e-test-report-v1.0.1.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This report records the real release validation results for `DrawioPpt v1.0.1`. The focus is both the real release pipeline and the SVG vector-text regression that matters most for this version.

## Test Environment

- OS: `Microsoft Windows 10 Pro 10.0.19045 64-bit`
- PowerPoint: `Microsoft PowerPoint Desktop x64`
- .NET: `.NET Framework 4.8`
- WebView2 Runtime: installed
- draw.io Desktop: `C:\Program Files\draw.io\draw.io.exe`
- Release version: `v1.0.1`

## Commands Actually Executed

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\restore-packages.ps1
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe .\DrawioPpt.sln /t:Rebuild /p:Configuration=Release /p:Platform=x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.1 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.1\package\scripts\install-release.ps1 -InstallRoot C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1\scripts\url-editor-smoke.ps1 -SkipBuild
```

These steps were the real release acceptance path for this version and covered build, packaging, installation, host/editor connectivity, and the SVG vector-text regression.

## Additional Regression Coverage

- Text fallback inspection on SVG exported by draw.io Desktop CLI
- SVG sanitization output inspection inside the add-in
- Inspection of the final SVG stored inside the saved `pptx` after real PowerPoint insertion

## Key Observations

- Original exported SVG:
  - `OriginalForeignObject=True`
  - `OriginalEmbeddedTextPng=True`
- Sanitized SVG:
  - `SanitizedForeignObject=False`
  - `SanitizedEmbeddedTextPng=False`
  - `SanitizedHasTspan=True`
- SVG stored inside the saved `pptx` after PowerPoint insertion:
  - `StoredSvgContainsForeignObject=False`
  - `StoredSvgContainsEmbeddedTextPng=False`
  - `StoredSvgHasTspan=True`

## Summary Table

- `ReleaseBuild`: `PASS`
  - `MSBuild Rebuild (Release|x64)` completed with `0 warning / 0 error`
  - `DrawioPpt.Core.dll = 1.0.1.0`
  - `DrawioPpt.PowerPointAddIn.dll = 1.0.1.0`
- `ReleasePackage`: `PASS`
  - `artifacts\releases\v1.0.1\package` was generated successfully
  - `artifacts\releases\v1.0.1\DrawioPpt-v1.0.1.zip` was generated successfully
- `InstallRelease`: `PASS`
  - The release package was installed to `C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1`
  - Registry `CodeBase` was updated to the installed `DrawioPpt.PowerPointAddIn.dll`
- `InstalledUrlSmoke`: `PASS`
  - The packaged `url-editor-smoke.ps1 -SkipBuild` returned `0`
  - The plug-in log recorded `configure -> init -> load -> save -> export -> DialogResult=OK`
- `SvgVectorTextRegression`: `PASS`
  - The original draw.io export sample contained `foreignObject + data:image/png`
  - Both the sanitized SVG and the SVG stored inside the saved `pptx` no longer kept bitmap text fallback and retained `tspan`

## Artifacts and Logs

- Release package manifest: `artifacts\releases\v1.0.1\package\PACKAGE.txt`
- Installation verification root: `C:\Users\nabul\AppData\Local\Temp\DrawioPpt\release-verify-v1.0.1`
- URL smoke plug-in log: `C:\Users\nabul\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log`
- Log directory for this version: `artifacts\logs\v1.0.1\`

## Conclusion

The core fix for `v1.0.1` has already been validated through a real PowerPoint host path: draw.io bitmap text fallback is converted before insertion, and the SVG finally stored inside the `pptx` remains vector text.
