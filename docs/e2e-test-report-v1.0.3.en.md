# Full E2E Test Report (v1.0.3)

[English](./e2e-test-report-v1.0.3.en.md) | [中文](./e2e-test-report-v1.0.3.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This report records the real release validation results for `DrawioPpt v1.0.3`, focusing on release package generation, PowerPoint/Word installer flow, and the new draw.io Desktop XML tools configuration documentation.

## 1. Environment

- OS: Windows
- Office hosts: Microsoft PowerPoint Desktop, Microsoft Word Desktop
- Build configuration: `Release|x64`
- Registration: current-user COM Add-in registration
- Recommended external editor: `draw.io Desktop v30.0.4 XML tools build`

## 2. Commands

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.3 -Configuration Release -Platform x64 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.3\package\scripts\install-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV103-2814525ef98e4c5a9e03c7cdb3e30c3e
powershell.exe -ExecutionPolicy Bypass -File .\artifacts\releases\v1.0.3\package\scripts\uninstall-release.ps1 -InstallRoot %TEMP%\DrawioPptReleaseV103-2814525ef98e4c5a9e03c7cdb3e30c3e
```

## 3. Build and Package Results

- Release build passed.
- Build result: 0 warnings and 0 errors.
- Release folder generated: `artifacts\releases\v1.0.3\package`
- Release zip generated: `artifacts\releases\v1.0.3\DrawioPpt-v1.0.3.zip`
- Package inspection confirmed no PDB, log, or screenshot assets.

## 4. Install Validation

- The release package installer registered the PowerPoint COM Add-in successfully.
- The release package installer registered the Word COM Add-in successfully.
- The release package uninstaller unregistered the PowerPoint COM Add-in successfully.
- The release package uninstaller unregistered the Word COM Add-in successfully.

After the temporary install validation, the local default PowerPoint/Word add-in registration was restored.

## 5. draw.io Desktop XML Tools Documentation Check

This version documents the recommended editor:

- Version: `draw.io Desktop v30.0.4 XML tools build`
- Asset: `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip`
- Runtime path: `win-unpacked\draw.io.exe`
- SHA256: `B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D`

The docs explain the XML tool button capabilities:

- Paste draw.io XML source from the clipboard and render it as a diagram.
- Copy the current canvas as draw.io XML source.

## 6. Known Limitation

This release mainly delivers the Word add-in, documentation closure, and desktop editor configuration guidance. It does not change the URL editor save protocol or SVG generation logic.
