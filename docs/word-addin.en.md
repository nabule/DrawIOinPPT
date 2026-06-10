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

## 3. Data Strategy

Word does not provide a direct equivalent to PowerPoint `Shape.Tags`, so the Word add-in uses:

- Picture `AlternativeText`: compressed Draw.io metadata envelope for quick detection and fallback compatibility.
- `Document.CustomXMLParts`: the full metadata envelope, so source XML moves with the `.docx`.
- Picture `Title`: a readable diagram name.

Reads use `AlternativeText` to resolve the `diagramId`, then resolve the newest full envelope from `Document.CustomXMLParts`. If only `AlternativeText` is available, the add-in rehydrates the document-level CustomXMLParts entry.

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
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Debug
```

Coverage:

- Debug solution build passes.
- A real Word COM session opens a `.docx`.
- URL mode creates a Draw.io diagram and writes SVG/XML back.
- The document is saved, reopened, edited again, and written back again.
- `Document.CustomXMLParts` persists.
- Word can load the COM add-in through `COMAddIns.Item("Greensoft.DrawioWordAddIn")`, with `Connect=True`.

`word-url-e2e.ps1` does not force-close user Word processes. If Word is already running, it aborts to avoid affecting unsaved documents.
