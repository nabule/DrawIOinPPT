# Installation and Local Debugging Guide

[English](./installation.en.md) | [中文](./installation.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Prerequisites

- Windows
- Microsoft PowerPoint Desktop x64
- Microsoft Word Desktop x64 when using the Word add-in
- .NET Framework 4.8 Developer Pack
- WebView2 Runtime
- Optional but recommended: [draw.io Desktop v30.0.4 XML tools build](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1)

Recommended first step:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\dev-check.ps1
```

## 2. Installation for End Users

If you received a release package such as `DrawioPpt-v1.0.10.zip`, the recommended installation flow is:

1. Extract the zip file to a local directory.
2. Make sure PowerPoint and Word are not currently running.
3. Double-click `install.cmd`.
4. Wait for the script to report success.
5. Open PowerPoint or Word and confirm that the `Draw.io` tab appears on the Ribbon.

Successful installation output includes `BuildId` in the form `v1.0.10+<git-short-hash>`. The same value is written to `BuildInfo.txt` / `PACKAGE.txt` in the package root and installed root, and installation verification reads the same text from the actual Word and PowerPoint add-ins.

On machines where an older build is already installed, upgrade by running `install.cmd` from the new release package. The installer stages the new package payload under `%TEMP%` before cleaning the old installation directory, so it can upgrade even when the package is extracted into the current install root or a child folder. Word and PowerPoint still need to be closed first; if Office or WebView2 keeps an old DLL locked, close Office and retry, or reboot Windows and run the new package's `install.cmd` again. On failure, the command window keeps the original PowerShell error and prints the exit code; follow that error before rerunning the new package.

Core installation files included in the release package:

- `install.cmd`
- `uninstall.cmd`
- `scripts\install-release.ps1`
- `scripts\uninstall-release.ps1`
- `scripts\register-office-addins.ps1`
- `scripts\unregister-office-addins.ps1`
- `scripts\verify-office-install.ps1`
- `BuildInfo.txt`
- `PACKAGE.txt`
- `scripts\register-addin.ps1`
- `scripts\unregister-addin.ps1`
- `scripts\register-word-addin.ps1`
- `scripts\unregister-word-addin.ps1`
- `scripts\install-release-upgrade-safety-test.ps1`
- `scripts\install-release-encoding-safety-test.ps1`
- `docs\release-notes-v1.0.10.en.md`
- `docs\e2e-test-report-v1.0.10.en.md`
- `docs\release-evidence-v1.0.10.en.md`

Default installation path:

```text
C:\Users\<your-username>\AppData\Local\Greensoft\DrawioPpt
```

How to uninstall:

1. Close PowerPoint.
2. Close Word.
3. Double-click `uninstall.cmd`.
4. Confirm that PowerPoint and Word no longer load the add-in automatically.

## 3. Installation for Repository-Based Development

If you are developing and debugging directly from this repository, build the solution first and then register the current build output.

## 4. Build

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

Default outputs:

- `src\DrawioPpt.Core\bin\x64\Debug\DrawioPpt.Core.dll`
- `src\DrawioPpt.PowerPointAddIn\bin\x64\Debug\DrawioPpt.PowerPointAddIn.dll`
- `src\DrawioPpt.WordAddIn\bin\x64\Debug\DrawioPpt.WordAddIn.dll`

The build script also writes `BuildInfo.txt` into all three output directories. At runtime the add-ins read this file first, so repository-based registration and release-package installation both show the same version and Git short hash in the settings window and Ribbon information area.

## 5. Register the Add-In

The current scripts use per-user COM registration, so administrator privileges are not required.

Register both PowerPoint and Word:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-office-addins.ps1
```

Notes:

- In a repository checkout, the script automatically resolves `src\DrawioPpt.PowerPointAddIn\bin\...` and `src\DrawioPpt.WordAddIn\bin\...`
- In a release package, the script automatically resolves `bin\DrawioPpt.PowerPointAddIn.dll` and `bin\DrawioPpt.WordAddIn.dll` next to the package root
- `install.cmd` / `scripts\install-release.ps1` also call this unified registration entry, so end-user installation covers both PowerPoint and Word

Unregister both PowerPoint and Word:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\unregister-office-addins.ps1
```

For single-host diagnostics, use the host-specific scripts:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-addin.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\register-word-addin.ps1
```

Post-install acceptance check:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

This script checks the PowerPoint/Word add-in DLLs, current-user COM registration, `CodeBase` targets, Office disabled items, and actual `COMAddIns` loading in Word and PowerPoint.

## 6. Recommended First-Run Setup

1. Open PowerPoint.
2. Confirm that `Greensoft.DrawioPptAddIn` is loaded; when using Word, also confirm `Greensoft.DrawioWordAddIn`.
3. Open the add-in settings window.
4. Configure the following as needed:
   - `Desktop Path`
   - `Editor URL`
   - `Auto update on save`
   - `Keep sidecar file next to PPT`

## 7. Recommended draw.io Desktop XML Tools Build

For desktop mode, use [draw.io Desktop v30.0.4 XML tools build](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1). It is based on official `v30.0.4` and adds two colored XML icon buttons to the second toolbar row:

- Paste draw.io XML source from the clipboard and render it as a diagram.
- Copy the current canvas as draw.io XML source.

Installation and configuration:

1. Download `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip` from the release page.
2. Verify SHA256:

   ```text
   B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D
   ```

3. Extract the zip to a stable folder, for example `D:\Tools\draw.io-30.0.4-xml-tools.1\`.
4. Run `win-unpacked\draw.io.exe`. This is an unsigned Windows x64 unpacked build; do not copy only the single exe.
5. In DrawioPpt settings, set `Editor Mode` to `Desktop`.
6. Set `Desktop Path` to the extracted `win-unpacked\draw.io.exe`, for example:

   ```text
   D:\Tools\draw.io-30.0.4-xml-tools.1\win-unpacked\draw.io.exe
   ```

When manually diagnosing source data between the add-in and draw.io Desktop, this build's XML copy/paste buttons are more convenient than the stock desktop build.

## 8. Recommended Post-Install Validation Order

Use this order to confirm that installation succeeded:

1. The add-in loads correctly.
2. The settings window opens.
3. The settings window footer and Ribbon information area show `版本：v1.0.10+<git-short-hash>`.
4. `Desktop` mode can detect a local `draw.io.exe`.
5. `Test URL` succeeds in `Url` mode.
6. You can create a new diagram and write it back successfully.

If you want to run the full automated acceptance flow:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\full-e2e-test.ps1
```

Full test report:

- [v1.0.10 E2E test report](./e2e-test-report-v1.0.10.en.md)

## 9. Desktop Mode Validation

1. Set `Editor Mode` to `Desktop` in settings.
2. Confirm that `Desktop Path` points to `draw.io.exe` or `diagrams.net.exe`.
3. Create a new diagram.
4. Save in the desktop editor and confirm that the SVG in PowerPoint refreshes automatically.

Additional notes:

- In `v1.0.5`, `simpleLabels` is enabled automatically only in `Url` mode for better MS Office SVG compatibility.
- In `Desktop` mode, the add-in launches the external desktop draw.io application and cannot rewrite that application's editor settings for you.
- Prefer the `win-unpacked\draw.io.exe` extracted from `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip` when you need manual XML copy/paste diagnostics.
- If you want clearer SVG text scaling in PowerPoint for desktop exports, configure draw.io Desktop manually:
  - `Extras > Configuration`
  - Enter `{ "simpleLabels": true }`
  - Click `Apply`

## 10. URL Mode Validation

1. Set `Editor Mode` to `Url` in settings.
2. Fill in `Editor URL`.
3. Click `Test URL` in the settings window first.
4. After the test passes, create or edit a diagram in PowerPoint.

If you want to run the repository URL smoke test:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\url-editor-smoke.ps1
```

From `v1.0.7`, the URL editor keeps its WebView2 profile in the current-user directory:

```text
%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2
```

Word and PowerPoint therefore do not need, and must not be granted, write access to the Office installation directory. If you previously saw “Cannot initialize URL editor” with access denied, install `v1.0.7`, fully close and reopen Office, then run `Test URL`. The following real Word-host regression can validate the release DLL from the repository. Close Word first; the script temporarily replaces and finally restores the current user's Word add-in registration and settings:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release
```

## 11. Log Location

Default log file:

```text
C:\Users\<your-username>\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

Useful for diagnosing:

- URL initialization timeout
- URL export timeout
- Host write-back failures
- Storage migration and orphan cleanup issues

## 12. Frequently Asked Questions

### An already-installed machine cannot upgrade from the installer

This is usually not a Word or PowerPoint registration problem. The old installer sequence deleted the default install directory before copying from the release package. If the user had extracted the new package into the existing install directory or one of its child folders, the cleanup step deleted the package itself, and the following registration step could not find the scripts or DLLs.

Use the fixed package and run `install.cmd` from that new package. Do not run an old `install.cmd` left in the previous installation directory. The fixed installer first copies the package payload listed in `PACKAGE.txt` into a temporary staging directory, then cleans the old install root and installs from staging, so it can replace the old version without bringing stale files back.

### `install.cmd` shows garbled text or a PowerShell parser error

This means Windows PowerShell decoded a PowerShell script in an old release package with the wrong encoding, before installation began. Download `v1.0.10` or later and do not rerun that old package. The new package uses ASCII-only PowerShell scripts, and `install.cmd` prints both the original error and its exit code on failure.

### The URL window opens, but save is never called back

- Click `Test URL` in settings first.
- Check the `init/save/export` events in the log.
- Confirm that the target URL supports the draw.io embed protocol.

### The URL editor cannot initialize or reports access denied

- Confirm that `v1.0.7` or later is installed, then fully close and reopen Word/PowerPoint.
- Verify that `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2` can be created and written by the current user; do not alter permissions under `C:\Program Files\Microsoft Office`.
- If the folder is writable but initialization still fails, inspect the `UrlEditor` log entries and confirm that the WebView2 Runtime is installed.

### Desktop mode opens the editor, but the shape does not refresh

- Confirm that the `.drawio` file was actually saved.
- Confirm that `Auto update on save` is enabled.
- Check whether the log contains a refresh exception.

### Can I still edit the diagram after moving the PPT file

- Yes. The source XML is stored in `Presentation.CustomXMLParts`.
- In `v1.0.5`, the add-in rebuilds the sidecar against the current PPT path on the next edit. If you move the sidecar with the deck, desktop-mode refresh is smoother.

### Why do SVG labels in URL mode look sharper when scaled

- In `v1.0.5`, the add-in automatically sends `{ "simpleLabels": true }` to the embedded URL editor.
- This follows the optimization recommended by draw.io for MS Office SVG scaling compatibility.

### The install script says PowerPoint is still running

- Close all PowerPoint windows and try again.
- If `POWERPNT.EXE` is still running in the background, end that process first and then retry installation or uninstall.

### The add-in does not appear after installation

- Reopen PowerPoint first.
- Then verify that the current-user registration entries were written.
- If needed, run `scripts\register-office-addins.ps1` again.
