# Full E2E Test Report (v1.0.5)

[English](./e2e-test-report-v1.0.5.en.md) | [中文](./e2e-test-report-v1.0.5.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This report records the real validation results for `DrawioPpt v1.0.5`, focusing on the default install path, Word registration, real Word COM loading, and the new install acceptance script.

## 1. Test Environment

- Operating system: Windows
- Office hosts: Microsoft PowerPoint Desktop and Microsoft Word Desktop
- Build configuration: `Release|x64`
- Registration mode: current-user COM Add-in registration
- Default install directory: `%LOCALAPPDATA%\Greensoft\DrawioPpt`

## 2. Reproduced Issue

At the start of the re-check, the default install directory still contained stale package contents and these registration entries were missing:

- `HKCU\Software\Microsoft\Office\PowerPoint\Addins\Greensoft.DrawioPptAddIn`
- `HKCU\Software\Microsoft\Office\Word\Addins\Greensoft.DrawioWordAddIn`
- `HKCU\Software\Classes\Greensoft.DrawioWordAddIn`
- `HKCU\Software\Classes\CLSID\{F10C5C83-0D86-4C81-A0B8-7E8FE9D31D8D}`

This matches the user's observation that Word was not installed. Further reproduction confirmed that the old `scripts\word-addin-load-check.ps1` unregistered Word at the end of the test, directly causing Word registration to disappear after checks.

## 3. Commands

```powershell
.\artifacts\releases\v1.0.4\package\install.cmd
powershell -ExecutionPolicy Bypass -File .\scripts\verify-office-install.ps1
gh release download v1.0.4 --repo nabule/DrawIOinPPT --pattern "DrawioPpt-v1.0.4.zip"
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\package-release.ps1 -Version v1.0.5 -Configuration Release -Platform x64 -SkipBuild
```

Before final release, also run:

```powershell
.\artifacts\releases\v1.0.5\package\install.cmd
powershell -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Release
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

## 4. Key Results

- The GitHub `v1.0.4` release asset SHA256 matched the local package: `9A253FAD59689D76F43002BA1CCBA7C841A78EAAD4BA8164ABE0BE9EB4066197`.
- The downloaded `v1.0.4` package included `register-office-addins.ps1`.
- Running `install.cmd` from the downloaded GitHub package wrote the Word registration entry.
- `Word.COMAddIns.Item("Greensoft.DrawioWordAddIn")` was found and connected with `Connect=True`.
- The old `word-addin-load-check.ps1` removed Word registration, so the regression check failed first; after the fix, the same check passed and `CodeBase` stayed pointed at the default install directory.
- The existing Word `Resiliency\DisabledItems` entry was not DrawioWord; it was Microsoft Word Previewer.
- After adding `verify-office-install.ps1`, default install directory acceptance passed.

Acceptance script output:

```text
PowerPoint registration verified: Greensoft.DrawioPptAddIn
Word registration verified: Greensoft.DrawioWordAddIn
Word.Application COM load verified: Greensoft.DrawioWordAddIn
PowerPoint.Application COM load verified: Greensoft.DrawioPptAddIn
DrawioPpt Office installation verified.
```

## 5. Conclusion

`v1.0.5` fixes the load-check script that unregistered Word, and makes default install path plus real Word COM loading part of the release gate. If Word is not registered, `CodeBase` points to the wrong file, Office disabled the add-in, or `COMAddIns` cannot connect, the acceptance script fails directly.
