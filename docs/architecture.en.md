# Architecture

[English](./architecture.en.md) | [中文](./architecture.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Goals

The goal of this project is to give Draw.io diagrams inside Office documents the following capabilities. PowerPoint Desktop is already covered, and this branch adds a Word Desktop host:

- Insert the original SVG directly in PowerPoint for vector quality; use a 1240px aspect-preserving PNG preview in Word to reduce native layout cost for complex SVG
- Use the existing PowerPoint edit entry point; in Word, re-enter Draw.io only after selecting a picture and clicking an explicit command
- Configure the editor as either a local desktop program or a web URL
- Keep the original Draw.io XML moving with the `pptx` / `docx` file as much as possible

## 2. Technical Direction

This project chooses **native Office COM Add-ins**, currently with PowerPoint and Word hosts, for these reasons:

- It can reliably listen to Office Desktop selection events
- It can launch local processes directly, which fits desktop draw.io integration
- It handles local paths, temporary files, and later write-back more naturally

WPS is out of scope for now, and Office Web Add-ins are not the primary direction.

## 3. Module Breakdown

### 3.1 `DrawioPpt.Core`

Responsibilities:

- Define the add-in settings model
- Define the Draw.io metadata envelope model
- Provide metadata serialization and compression
- Generate sidecar file paths

The shared settings model includes editor mode, desktop editor path, URL editor address, Office-compatible SVG labels, auto-open-on-selection, auto-refresh-after-save, sidecar retention, and whether to show diagram information when creating or editing. Auto-open-on-selection is used only by PowerPoint. Word does not subscribe to selection changes or expose auto-open, keeping the drag hot path free of add-in work.

### 3.2 `DrawioPpt.PowerPointAddIn`

Responsibilities:

- COM add-in entry point and registration
- Ribbon command exposure
- Selected-shape monitoring
- PowerPoint object model interaction
- External-editor launch and write-back orchestration

### 3.3 `DrawioPpt.WordAddIn`

Responsibilities:

- Word COM add-in entry point and registration
- Ribbon command exposure
- Live picture recognition triggered by explicit Word commands
- Word `InlineShape` / floating `Shape` object model interaction
- Word document-level `CustomXMLParts` storage and cleanup
- External-editor launch and write-back orchestration

The current Word host reuses public editor, SVG-export, and SVG-sanitization services from the PowerPoint assembly. These shared services can later be extracted into a dedicated `OfficeShared` project. Display strategy is host-specific: PowerPoint inserts the original SVG, scales it proportionally with `contain` inside the slide bounds, and centers it without rewriting the `viewBox` or padding the canvas. Word renders an aspect-preserving 1240px-wide PNG from the Draw.io source, inserts it as a floating `Shape`, and fixes wrapping to `wdWrapFront`. If draw.io Desktop export fails, Word generates the lightweight PNG placeholder from a 1240px baseline while preserving the source aspect ratio; extreme portrait diagrams remain protected by a 1200px height cap. It does not put a complex SVG back into Word. Complete source remains in document metadata and is still editable after selection plus an explicit command.

Word does not subscribe to `WindowSelectionChange`, and startup does not read the current selection, picture properties, or `AlternativeText`. Selecting, dragging, resizing, and repositioning therefore use Word's native path without add-in metadata detection, Ribbon-state invalidation, document scanning, or auto-open work. Word Ribbon commands remain enabled: the user selects a picture and then clicks Re-edit, Refresh, Bind, or Clear Binding, at which point the command reads and validates the live selection. Double-click remains an explicit editing action and may read the current picture.

### 3.4 Installation and Registration Scripts

Responsibilities:

- `scripts\register-office-addins.ps1` is the unified registration entry for repository checkouts and release packages. It registers both current-user COM Add-ins: PowerPoint and Word.
- `scripts\unregister-office-addins.ps1` is the unified unregister entry and removes both PowerPoint and Word registration entries.
- `scripts\verify-office-install.ps1` is the installation acceptance entry. It checks installed files, current-user COM registration, `CodeBase` targets, Office disabled items, and optional Word/PowerPoint COM loading and version-automation callbacks.
- `scripts\register-addin.ps1`, `scripts\register-word-addin.ps1`, and their unregister counterparts remain available for single-host diagnostics.
- The release package's `install.cmd` calls `scripts\install-release.ps1`; after copying files, the installer calls the unified registration entry and immediately runs non-interactive registration acceptance checks so Word is not missed when PowerPoint is installed.
- Upgrade installation follows a stage-clean-copy order: `install-release.ps1` first copies the manifest-listed package payload into `%TEMP%`, then removes the old install directory and installs from staging. This avoids deleting the installer source when the new package is inside the old install root, and avoids bringing stale old files back.
- Office calls the Ribbon version label through `GetVersionSummary` on each `Connect` COM entry. The entry forwards to its `RibbonController` and shared `BuildInfo`, so the BuildId shown by Word and PowerPoint matches the package `BuildInfo.txt`. Installation acceptance reads the same host version text through each host's `COMAddIn.Object` automation object; it is separate from the Ribbon entry but reaches the same `AddInHost`.

## 4. Data Storage Strategy

### Initial Version Strategy

To reach a usable MVP quickly, the initial version uses two storage layers:

PowerPoint:

- Shape identity: `Shape.Tags["DRAWIO_PPT_ID"]`
- Shape metadata envelope: `Shape.AlternativeText`

Word:

- Picture metadata envelope: `InlineShape/Shape.AlternativeText`
- Diagram display name: `InlineShape/Shape.Title`

`AlternativeText` stores a custom XML envelope defined by the add-in, containing:

- diagram id
- display name
- editor mode
- editor target
- sidecar path
- updated time
- original draw.io XML

The original draw.io XML is compressed using `gzip + base64` first to reduce size.

### Current Embedding Model

PowerPoint has now been enhanced to `Presentation.CustomXMLParts + Shape.Tags + Shape.AlternativeText`; Word uses `Document.CustomXMLParts + InlineShape/Shape.AlternativeText`.

This means the Draw.io source XML is already embedded inside the `.pptx` / `.docx` file, instead of relying only on an external `.drawio` file. The responsibilities are split as follows:

- Document-level `CustomXMLParts` are the primary store. Both PowerPoint and Word keep the full add-in envelope there, including `formatVersion`, `diagramId`, diagram name, editor mode, editor target, sidecar path, update time, and compressed `DrawioXml`.
- `DrawioXml` is stored in the envelope's `drawioXml` node using `gzip + base64`, reducing Office file size and avoiding large raw XML payloads in shape properties.
- PowerPoint shapes point back to the document-level XML part through `Shape.Tags["DRAWIO_PPT_ID"]` and `Shape.Tags["DRAWIO_PPT_PART_ID"]`.
- Word has no direct equivalent to `Shape.Tags`. On a normal write, `InlineShape/Shape.AlternativeText` holds a lightweight envelope without `DrawioXml`: `formatVersion`, `diagramId`, diagram name, editor mode, editor target, sidecar path, and update time. The add-in uses that reference to load the full envelope from `Document.CustomXMLParts`.
- If Word `CustomXMLParts.Upsert` fails, the add-in leaves the full envelope in picture `AlternativeText` so that the save does not lose source data. This failure fallback does not change the PowerPoint storage strategy.
- A legacy Word document may have its full envelope only in `AlternativeText`. On first read, the add-in migrates the full data into `Document.CustomXMLParts` and, after a successful write, rewrites the picture with a lightweight reference. Repeated reads do not keep creating new parts.
- The visible SVG also receives draw.io `content` metadata as a third recovery layer. This is not the primary store, but it helps when a single SVG is exported or inspected separately.
- The sidecar `.drawio` file is now closer to an editing cache and manual backup. It is still useful when draw.io Desktop needs a real file path, while the in-document XML is the primary source for cross-machine document movement.

The current implementation does not embed `.drawio` as an OLE object or `EmbeddedPackagePart` attachment. Avoiding that path reduces Office security prompts, file-association dependency, and host-specific differences; the tradeoff is that the add-in must maintain the index from visible shape/picture back to the document-level XML part.

### Known Boundaries

- Document-level `CustomXMLParts` travel when the whole `.pptx` / `.docx` file is saved and moved, but Office does not guarantee that copying one shape into another document also copies its matching document-level XML part.
- For Word, the normal lightweight `AlternativeText` does not contain Draw.io XML. Editing source may therefore be unrecoverable after copying only one picture into another document. Copy the complete `.docx`, or save/retain the sidecar `.drawio` or Draw.io XML first. Only legacy pictures and primary-store failure fallbacks may still contain the full payload in `AlternativeText`.
- Document Inspector, enterprise DLP, saving to older formats, exporting to PDF, or third-party Office-compatible software may remove or ignore custom XML parts.
- `AlternativeText` is a visible accessibility/description property and should not be treated as the long-term high-capacity primary store. Word uses it as a normal recognition reference and as a full-payload fallback only during a failed primary-store write or legacy migration.
- Clearing binding removes add-in metadata from the visible shape/picture and cleans unreferenced document-level XML parts, but it does not delete the visible picture itself.

## 5. Editing Workflow

### 5.1 Create

1. Click `New Diagram` on the Ribbon
2. Launch the external draw.io editor
3. The user saves the `.drawio` file
4. The add-in exports SVG; the normal Word path additionally renders a 1240px aspect-preserving PNG from the Draw.io source
5. PowerPoint inserts the original SVG with centered `contain` sizing; Word inserts the PNG at the cursor as a floating `wdWrapFront` picture
6. The add-in writes the metadata envelope and document-level `CustomXMLParts`
7. If `ShowDiagramInfoDialog` is enabled, it shows the diagram name, Diagram ID, editor mode, and `.drawio` working file path

### 5.2 Edit

1. The user selects a shape or picture
2. PowerPoint may update status through its existing selection event; Word selection itself performs no add-in work
3. The Word user clicks Re-edit (or double-clicks the picture), while PowerPoint uses its existing edit entry point
4. After that explicit action, the add-in identifies the object and reads the metadata envelope
5. It opens the local editor or URL editor
6. The user saves
7. The add-in regenerates displayed content: PowerPoint still replaces with the original SVG, while Word regenerates the 1240px aspect-preserving PNG and replaces the floating picture
8. If `ShowDiagramInfoDialog` is enabled, it shows the current diagram information when entering the edit flow

### 5.3 WebView2 User Data Folder for the URL Editor

Before creating `WebView2`, URL mode explicitly creates a `CoreWebView2Environment` and fixes the browser user-data directory at:

```text
%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2
```

The current user creates this directory and PowerPoint and Word share it. This prevents WebView2 from applying its default profile location next to `POWERPNT.EXE` or `WINWORD.EXE`; the latter is normally under the protected Office installation directory and can fail with `E_ACCESSDENIED`. The directory holds only the WebView2 browser profile, not Draw.io source XML; diagram source remains managed by Office `CustomXMLParts` and picture metadata.

## 6. Update Strategy

The first stage updates diagrams by replacing the displayed object, prioritizing an end-to-end working flow.

The following data is preserved:

- position
- size
- rotation
- z-order

On create, PowerPoint applies centered `contain` sizing to the original SVG within the slide bounds. On replacement it retains the existing width and corrects height from the SVG source ratio without expanding the SVG `viewBox`, so no internal blank space is introduced. Word uses a 1240px aspect-preserving PNG for create, Re-edit, and Refresh. The fallback placeholder is generated from a 1240px baseline while preserving the source aspect ratio, with a 1200px height cap for extreme portrait diagrams. The picture is always converted to a floating `Shape` with `wdWrapFront`. Both horizontal and vertical diagrams use `contain` within the document bounds; the vertical regression ratio is `0.25`. Replacement is transactional: create and fully configure the new picture before deleting the original; if creation or configuration fails, remove the new object and preserve the original. Temporary preview XML deletion uses bounded retries and delays, failed directories receive deferred background cleanup, and provider startup also removes stale directories.

If later validation shows that animations, hyperlinks, or complex formatting are lost too easily during replacement, the enhancement path is:

- Replace the internal picture part inside the PPTX directly to reduce object reconstruction

## 7. Risks and Responses

### Risk 1: `AlternativeText` capacity and compatibility

Response:

- Keep only a lightweight reference in Word `AlternativeText` during normal operation, with the full envelope in `Document.CustomXMLParts`.
- Retain a full picture fallback when `CustomXMLParts` cannot be written, and migrate legacy full metadata idempotently on read.
- Document the single-picture cross-document copy tradeoff and recommend copying the complete document or retaining sidecar/Draw.io XML.

### Risk 2: Detecting saves from the external editor

Response:

- In local mode, prioritize write-back after process exit first
- Add file monitoring and throttling in the second stage

### Risk 3: SVG replacement fidelity

Response:

- Preserve position and size first
- Then address animation and complex-style issues with targeted fixes

### Risk 4: Word flowing layout and floating-picture anchors

Response:

- Prioritize the `InlineShape` picture workflow for the MVP
- Preserve anchor, size, wrapping mode, and relative position when replacing floating `Shape` pictures
- Validate save, reopen, and edit-again behavior with a real Word E2E test
