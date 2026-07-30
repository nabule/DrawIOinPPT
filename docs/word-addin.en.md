# Word Add-in Design and Usage

[English](./word-addin.en.md) | [中文](./word-addin.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Goal

The Word add-in brings the current PowerPoint Draw.io workflow to Word Desktop:

- Insert a Draw.io diagram at the current cursor position.
- Render a 1240px-wide PNG preview from the Draw.io source while preserving its aspect ratio.
- Display it as a floating `wdWrapFront` picture; editing starts only after the user selects the picture and clicks a command.
- Write the updated PNG preview and Draw.io XML back to the original picture after save.
- Store source XML in `Document.CustomXMLParts`, with picture `AlternativeText` carrying the lightweight reference and failure fallback.
- Keep supporting both local draw.io / diagrams.net Desktop and URL editor modes.

## 2. Current Scope

The new project is `src/DrawioPpt.WordAddIn/`, an independent Word COM Add-in host. To keep this first implementation focused, the Word host reuses the existing shared editor and SVG services from `DrawioPpt.PowerPointAddIn`, including:

- Desktop draw.io launch and SVG export.
- URL-mode `WebView2 + diagrams.net embed` editor.
- Settings UI, logging, user notifications, and the version label in the settings footer.
- SVG export and sanitization, 1240px PNG preview generation, and draw.io `content` metadata embedding.

Word-specific code handles:

- Word COM Add-in registration and Ribbon callbacks.
- Explicit Word double-click handling, without monitoring ordinary selection changes.
- `InlineShape` / floating `Shape` picture detection.
- 1240px aspect-preserving PNG floating-picture insertion and replacement.

Word uses a display path separate from PowerPoint. PowerPoint continues to insert the original SVG directly. The normal Word path exports a 1240px-wide PNG from the Draw.io source, preserves source aspect ratio, explicitly requests `--border 0`, and converts the picture to a floating `Shape` with `wdWrapFront`. Both horizontal and vertical diagrams use `contain` within the document bounds; the vertical regression ratio is `0.25`. Re-edit and Refresh create and fully configure the new picture before deleting the original; failures remove the incomplete new object and preserve the original. If draw.io Desktop export is unavailable, the add-in generates a lightweight PNG placeholder from a 1240px baseline while preserving the source aspect ratio, with a 1200px height cap for extreme portrait diagrams, instead of falling back to complex SVG. Complete Draw.io XML remains in document-level primary storage, and the placeholder remains editable after selection plus Re-edit. Transparent-padding pixel inspection is not applicable to the actual opaque 24bpp PNG; acceptance records the independent SVG source-ratio comparison and `--border 0` without representing unsupported pixel inspection as passed.

The preview provider creates temporary Draw.io XML and PNG files only under `%TEMP%\DrawioPpt\word-preview`. Temporary XML deletion uses four bounded attempts separated by 50ms. If synchronous cleanup cannot finish, a deferred background retry is queued; provider startup also removes managed directories that are more than one hour old.
- `Document.CustomXMLParts` write/read/orphan cleanup.
- Sidecar `.drawio` relocation based on the Word document path.

### 2.1 Zero Selection Hot Path and Explicit Operations

The Word host does not subscribe to `WindowSelectionChange`, and startup does not read the current selection, picture properties, `AlternativeText`, or `Document.CustomXMLParts`. During ordinary selection, dragging, resizing, and repositioning, the add-in performs no metadata recognition, Ribbon-state refresh, orphan cleanup, or auto-open work; Word performs only its native picture interaction.

Re-edit, Refresh, Bind, and Clear Binding remain enabled. The user selects a picture and then clicks the command; only then does the add-in read the live Word selection and validate metadata. Ribbon information displays fixed select-then-click guidance rather than changing with selection. Double-click is also an explicit edit action and may open a managed picture. Ordinary selection, dragging, resizing, and repositioning perform no add-in metadata work. This release changes the normal display to a 1240px aspect-preserving PNG and `wdWrapFront` floating layout, but that does not establish that native Word mouse dragging passed acceptance.

## 3. Data Strategy

Word does not provide a direct equivalent to PowerPoint `Shape.Tags`, so the Word add-in uses:

- Picture `AlternativeText`: normally a lightweight envelope without `DrawioXml`, used for quick detection and primary-store lookup.
- `Document.CustomXMLParts`: the full metadata envelope, so source XML moves with the `.docx`.
- Picture `Title`: a readable diagram name.

The lightweight envelope retains `formatVersion`, `diagramId`, diagram name, editor mode, editor target, sidecar path, and update time, but not `DrawioXml`. Reads use it to resolve the `diagramId`, then load the newest full envelope from `Document.CustomXMLParts`.

If `CustomXMLParts.Upsert` fails, the add-in leaves the full envelope in picture `AlternativeText` so source data is not lost. If a legacy document has its full envelope only in `AlternativeText`, the first read migrates it into `Document.CustomXMLParts`. After successful migration the picture is rewritten with a lightweight reference, and the second read reuses the same primary-store part instead of migrating again.

### 3.1 In-Document Embedding for Word

The Word add-in also already embeds the Draw.io source XML inside the `.docx` file. The primary store is `Document.CustomXMLParts`, which contains the add-in envelope with:

- `diagramId`
- diagram display name
- current editor mode
- editor target, either a desktop executable path or URL
- sidecar `.drawio` path
- update time
- Draw.io XML compressed with `gzip + base64`

After selecting a picture and clicking Re-edit, the add-in resolves `diagramId` from picture `AlternativeText`, then finds the newest full envelope in `Document.CustomXMLParts`. After save, it updates the visible PNG preview, the lightweight picture reference, and the document-level XML. A full picture fallback is retained only if the primary-store write fails.

The sidecar `.drawio` file is also an editing cache and manual backup in the Word workflow, not the primary store. After moving the `.docx` to another machine, the add-in can still recover Draw.io XML from the document itself as long as `Document.CustomXMLParts` has not been stripped; if desktop draw.io is needed, the sidecar can be regenerated.

The current implementation does not embed `.drawio` files as OLE objects or Office attachments in Word. This avoids Office security prompts, external file-association dependency, and cross-machine opening differences. The compatibility tradeoff is that Office does not guarantee that copying one Word picture carries the source document's `CustomXMLParts`, while normal lightweight `AlternativeText` does not contain Draw.io XML. Editing source may therefore be unrecoverable in the destination document. Copy the complete `.docx`, or save/retain the sidecar `.drawio` or Draw.io XML first. Only legacy pictures and primary-store failure fallbacks may still carry the full payload in `AlternativeText`.

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
The Word and PowerPoint Ribbon info areas both show the current BuildId, formatted as `v1.0.9+<git-short-hash>`. The settings window footer shows the same value, which helps confirm whether Office loaded the DLLs from the current release package. Release-package verification directly calls both Office add-ins' version callbacks and compares them with `BuildInfo.txt`.

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

After developer registration, Word shows the `Draw.io` Ribbon:

- `New`: insert a new Draw.io diagram at the current cursor.
- `Re-edit`: edit the selected managed picture.
- `Refresh`: regenerate the picture from the bound `.drawio` working file.
- `Bind`: bring a normal picture under Draw.io management.
- `Clear Binding`: remove metadata while keeping the visible picture.
- `Settings`: reuse the existing editor settings.
- Info-area version: shows the current BuildId, for example `v1.0.9+<git-short-hash>`.

Word does not expose Auto Open, so selection and dragging remain free of add-in work. PowerPoint auto-open behavior is unchanged.

## 6. Validation

The current implementation has passed these real checks:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\build.ps1 -Configuration Release -Platform x64
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-complex-metadata-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-selection-event-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-url-e2e.ps1 -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e-safety-test.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-url-addin-host-e2e.ps1 -Configuration Release
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-svg-aspect-ratio-e2e.ps1 -Configuration Release -SkipBuild
powershell.exe -ExecutionPolicy Bypass -File .\scripts\word-addin-load-check.ps1 -Configuration Debug
```

Coverage:

- The Release x64 solution build passes with `0` warnings and `0` errors.
- A real Word COM session opens a `.docx`.
- URL mode creates a Draw.io diagram and writes the Word PNG preview/XML back.
- The document is saved, reopened, edited again, and written back again.
- `Document.CustomXMLParts` persists.
- Complex Draw.io XML remains in `CustomXMLParts`, while picture `AlternativeText` is a lightweight reference without `DrawioXml`; this remains true after save, close, and reopen.
- A full picture fallback persists when `CustomXMLParts` cannot be written, and the XML part ID remains stable on a second read after migrating legacy full metadata.
- Word can load the COM add-in through `COMAddIns.Item("Greensoft.DrawioWordAddIn")`, with `Connect=True`.
- `word-url-addin-host-e2e.ps1` loads the tested DLL inside real `WINWORD.EXE`, selects a managed picture, and invokes Re-edit through the add-in's explicit-command automation entry. It verifies `ExplicitEditAutomationAvailable=True`, `ExplicitEditCommandInvoked=True`, `configure -> init -> load -> save -> export` write-back, the per-user WebView2 folder, and no `E_ACCESSDENIED` in the log.
- `word-preview-image-provider-e2e.ps1` verifies the normal PNG signature, 1240px width, and source ratio, plus the 1240×620 aspect-preserving lightweight PNG placeholder produced for its 2:1 export-failure sample. It also verifies the safe temporary path and immediate source-file and preview-file cleanup, with no fallback to complex SVG. Bounded retries, deferred background cleanup, and startup cleanup are implementation mechanisms; this E2E does not fault-inject those mechanisms.
- `word-svg-aspect-ratio-e2e.ps1` verifies a 2:1 PNG remains 2:1 and stays a floating `wdWrapFront` picture after real Word insert and replacement; the vertical case reports `VerticalRatio=0.25` and `VerticalContained=True`. A failed replacement reports `FailedReplacementPreservedOriginal=True`, proving that the original is not deleted early.
- No residual `WINWORD.EXE` remains after the Word automation scripts finish.

`word-url-e2e.ps1` does not force-close user Word processes. If Word is already running, it aborts to avoid affecting unsaved documents. To isolate the host it creates, it temporarily disables automatic loading of the registered add-in and restores the original `LoadBehavior` in `finally`; do not start Word or run another Word automation task while it is running.

`word-url-addin-host-e2e.ps1` also requires Word to be closed, but temporarily registers the requested DLL and restores the Word add-in, COM class registrations, and settings when it finishes. It handles only Word processes reported by its helper with a matching start time: it first waits for bounded natural exit, then force-stops and polls again only if needed, rather than cleaning up every Word process that appears during the test. URL mode fixes the WebView2 profile at `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2`, so no Office installation-directory write access is required.
