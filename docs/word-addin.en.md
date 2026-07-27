# Word Add-in Design and Usage

[English](./word-addin.en.md) | [中文](./word-addin.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Goal

The Word add-in brings the current PowerPoint Draw.io workflow to Word Desktop:

- Insert a Draw.io diagram at the current cursor position.
- Display the diagram as an SVG picture where Word supports it.
- Re-open a selected managed picture in Draw.io.
- Write the updated SVG and Draw.io XML back to the original picture after save.
- Store source XML in `Document.CustomXMLParts`, with picture `AlternativeText` as a fallback.
- Keep supporting both local draw.io / diagrams.net Desktop and URL editor modes.

## 2. Current Scope

The new project is `src/DrawioPpt.WordAddIn/`, an independent Word COM Add-in host. To keep this first implementation focused, the Word host reuses the existing shared editor and SVG services from `DrawioPpt.PowerPointAddIn`, including:

- Desktop draw.io launch and SVG export.
- URL-mode `WebView2 + diagrams.net embed` editor.
- Settings UI, logging, and user notifications.
- SVG preview generation, SVG sanitization, and draw.io `content` metadata embedding.

Word-specific code handles:

- Word COM Add-in registration and Ribbon callbacks.
- Word selection monitoring and double-click handling.
- `InlineShape` / floating `Shape` picture detection.
- SVG picture insertion and replacement.
- `Document.CustomXMLParts` write/read/orphan cleanup.
- Sidecar `.drawio` relocation based on the Word document path.

### 2.1 Selection Synchronization and Picture Operation Performance

The Word host uses `WindowSelectionChange` to synchronize Ribbon state and reads the current selection once at startup. It no longer uses a 500 ms timer to repeatedly access Word COM and parse picture metadata, so that background work no longer periodically interrupts the UI thread while a picture is moved or resized.

Edit, Refresh, Bind, and Clear Binding are explicit operations and re-read the live Word selection on click. If Word does not immediately raise a selection event, users can still select the picture and invoke the relevant command. This release does not change SVG, Draw.io XML, picture format, wrapping, or the layout cost of a floating picture itself.

## 3. Data Strategy

Word does not provide a direct equivalent to PowerPoint `Shape.Tags`, so the Word add-in uses:

- Picture `AlternativeText`: compressed Draw.io metadata envelope for quick detection and fallback compatibility.
- `Document.CustomXMLParts`: the full metadata envelope, so source XML moves with the `.docx`.
- Picture `Title`: a readable diagram name.

Reads use `AlternativeText` to resolve the `diagramId`, then resolve the newest full envelope from `Document.CustomXMLParts`. If only `AlternativeText` is available, the add-in rehydrates the document-level CustomXMLParts entry.

### 3.1 In-Document Embedding for Word

The Word add-in also already embeds the Draw.io source XML inside the `.docx` file. The primary store is `Document.CustomXMLParts`, which contains the add-in envelope with:

- `diagramId`
- diagram display name
- current editor mode
- editor target, either a desktop executable path or URL
- sidecar `.drawio` path
- update time
- Draw.io XML compressed with `gzip + base64`

When re-editing a selected picture, the add-in first resolves `diagramId` from the picture `AlternativeText`, then finds the newest full envelope in `Document.CustomXMLParts`. After save, it updates the visible SVG picture, the fallback metadata on the picture, and the document-level XML.

The sidecar `.drawio` file is also an editing cache and manual backup in the Word workflow, not the primary store. After moving the `.docx` to another machine, the add-in can still recover Draw.io XML from the document itself as long as `Document.CustomXMLParts` has not been stripped; if desktop draw.io is needed, the sidecar can be regenerated.

The current implementation does not embed `.drawio` files as OLE objects or Office attachments in Word. This avoids Office security prompts, external file-association dependency, and cross-machine opening differences. The tradeoff is that copying a Word picture into another document may need to rely on `AlternativeText` or SVG `content` recovery because document-level `CustomXMLParts` are not guaranteed to travel with a single copied picture.

### 3.2 Recommended Desktop Editor

Word and PowerPoint currently share the same add-in settings. In desktop mode, set `Desktop Path` to the `win-unpacked\draw.io.exe` extracted from [draw.io Desktop v30.0.4 XML tools build](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1).

This build is based on official `v30.0.4` and adds two colored XML icon buttons to the second toolbar row:

- Paste draw.io XML source from the clipboard and render it as a diagram.
- Copy the current canvas as draw.io XML source.

This is especially useful for the Word add-in when you need to inspect the Draw.io XML embedded in a `.docx`, manually recover a diagram, or quickly put an XML snippet back into the desktop editor. Download `draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip` and fully extract it; do not copy only the single exe. SHA256:

```text
B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D
```

## 4. Release Package Installation

To install from a release package, extract `DrawioPpt-<version>.zip`, close any running PowerPoint and Word instances, then run this from the extracted folder:

```powershell
.\install.cmd
```

The installer copies files into the current user profile and registers both the PowerPoint and Word COM Add-ins. The default installation path is:

```text
%LOCALAPPDATA%\Greensoft\DrawioPpt
```

For a temporary verification install, pass an explicit install root:

```powershell
.\install.cmd -InstallRoot C:\tmp\DrawioPptInstallCheck
```

To uninstall, close PowerPoint and Word first, then run:

```powershell
.\uninstall.cmd
```

If an upgrade reports that a file such as `WebView2Loader.dll` is still in use, an old WebView2 child process has usually not exited yet. Close Office, the add-in settings window, and any URL editor window; if the file remains locked, reboot and rerun the installer. The installer does not terminate those processes automatically, to avoid affecting unrelated WebView2 applications.

After installation, Word shows the `Draw.io` Ribbon; the existing PowerPoint add-in remains available in PowerPoint.

## 5. Developer Build and Registration

Build:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1
```

Register the Word add-in:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\register-word-addin.ps1
```

Unregister the Word add-in:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\unregister-word-addin.ps1
```

After developer registration, Word shows the `Draw.io` Ribbon. Its entry points mirror the PowerPoint add-in:

- `New`: insert a new Draw.io diagram at the current cursor.
- `Re-edit`: edit the selected managed picture.
- `Refresh`: regenerate the picture from the bound `.drawio` working file.
- `Bind`: bring a normal picture under Draw.io management.
- `Clear Binding`: remove metadata while keeping the visible picture.
- `Auto Open`: automatically open the editor when selecting a managed picture.
- `Settings`: reuse the existing editor settings.

## 6. Validation

The current implementation has passed these real checks:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Debug -Platform x64
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -SkipBuild
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Debug
```

Coverage:

- Debug solution build passes.
- A real Word COM session opens a `.docx`.
- URL mode creates a Draw.io diagram and writes SVG/XML back.
- The document is saved, reopened, edited again, and written back again.
- `Document.CustomXMLParts` persists.
- Word can load the COM add-in through `COMAddIns.Item("Greensoft.DrawioWordAddIn")`, with `Connect=True`.
- `word-url-addin-host-e2e.ps1` loads the tested DLL inside real `WINWORD.EXE`, triggers the URL editor by selecting a managed picture, and verifies `configure -> init -> load -> save -> export` write-back, the per-user WebView2 folder, and no `E_ACCESSDENIED` in the log.

`word-url-e2e.ps1` does not force-close user Word processes. If Word is already running, it aborts to avoid affecting unsaved documents. To isolate the host it creates, it temporarily disables automatic loading of the registered add-in and restores the original `LoadBehavior` in `finally`; do not start Word or run another Word automation task while it is running.

`word-url-addin-host-e2e.ps1` also requires Word to be closed, but temporarily registers the requested DLL and restores the Word add-in, COM class registrations, and settings when it finishes. It stops only Word processes reported by its helper with a matching start time, rather than every Word process that appears during the test. URL mode fixes the WebView2 profile at `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`, so no Office installation-directory write access is required.
