# DrawioPpt

[English](./README.en.md) | [中文](./README.md)

DrawioPpt is a native Draw.io / diagrams.net add-in for Microsoft Office Desktop. It lets you insert, detect, re-edit, and refresh Draw.io diagrams directly in PowerPoint and Word, instead of maintaining diagrams as one-time screenshots.

Current public release: [DrawioPpt v1.0.8](https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.8)

Previous public release: [DrawioPpt v1.0.7](https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.7)

## Value

DrawioPpt is built for long-lived diagrams in Office documents:

- Diagrams are inserted as SVG for better visual quality than screenshots.
- A selected managed diagram can be reopened in Draw.io and written back to the same Office object after save.
- Draw.io source XML is embedded inside the `.pptx` / `.docx` file, so documents are not dependent only on neighboring `.drawio` files.
- PowerPoint and Word share the same editor settings.
- The PowerPoint and Word Ribbon info areas, plus the shared settings window, show the current BuildId so the loaded add-in build is easy to identify.
- Word does not monitor ordinary selection changes or read picture metadata at startup; selecting, dragging, resizing, and repositioning perform no add-in work.
- Both local draw.io Desktop and embedded URL editor modes are supported, covering offline, internal-network, and private diagrams.net deployments.

## Supported Office Hosts

### PowerPoint

- Insert a new Draw.io diagram into the current slide.
- Detect, re-edit, and refresh managed diagrams.
- Open the editor by double-clicking or selecting a managed shape.
- Store source data and references through `Presentation.CustomXMLParts + Shape.Tags + AlternativeText`.
- Relocate sidecar `.drawio` working files after save-as or folder moves.

### Word

- Insert a Draw.io diagram at the current cursor position.
- Re-edit, refresh, bind, and clear binding for selected managed pictures.
- Support both `InlineShape` and floating `Shape` pictures.
- Store source data and lightweight references through `Document.CustomXMLParts + AlternativeText`: the full Draw.io XML stays in the document-level primary store instead of being duplicated in the picture property.
- Share desktop editor, URL editor, and base settings with the PowerPoint add-in.
- Users select a picture first and then click Re-edit, Refresh, Bind, or Clear Binding; the explicit command reads the live selection once.

## Core Features

- `New`: create a Draw.io diagram and insert it into the current PPT/Word document.
- `Re-edit`: open the selected managed diagram and write changes back to the original object.
- `Refresh`: regenerate SVG from bound source data.
- `Bind`: bring a normal picture under add-in management.
- `Clear Binding`: remove add-in metadata while keeping the visible diagram.
- `Auto Open` (PowerPoint only): automatically enter edit flow when selecting a managed diagram. Word omits it to keep the drag hot path free of add-in overhead.
- `Settings`: configure desktop editor path, URL editor, sidecar behavior, and information dialogs.
- Version check: the Ribbon info area and settings footer show a BuildId in the form `v1.0.8+<git-short-hash>`.

## Source Data Storage

DrawioPpt uses layered persistence:

- Primary store: Office document-level `CustomXMLParts` containing compressed Draw.io XML.
- PowerPoint shape references: `Shape.Tags` stores `diagramId` and the document-level XML part reference.
- Word picture references: under normal operation, picture `AlternativeText` contains a lightweight envelope without `DrawioXml`: `formatVersion`, `diagramId`, name, editor mode/target, sidecar path, and update time. The full envelope remains in `Document.CustomXMLParts`.
- Word failure fallback: if writing `Document.CustomXMLParts` fails, picture `AlternativeText` retains the full envelope to avoid source-data loss. Legacy full picture metadata is migrated into the document-level primary store on first read.
- SVG fallback: generated SVG receives draw.io `content` metadata.
- sidecar `.drawio`: desktop editing cache, manual backup, and auto-refresh helper, not the only source of data.

This means Draw.io source data usually travels with the `.pptx` / `.docx` file itself. Document Inspector, enterprise DLP, saving to legacy formats, PDF export, or third-party Office-compatible software may remove or ignore custom XML.

## Recommended draw.io Desktop

Desktop mode works best with [draw.io Desktop v30.0.4 XML tools build](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1).

Recommended asset:

```text
draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip
```

After extraction, set `Desktop Path` to:

```text
win-unpacked\draw.io.exe
```

This build adds two XML buttons to the second draw.io toolbar row:

- Paste draw.io XML source from the clipboard and render it as a diagram.
- Copy the current canvas as draw.io XML source.

These tools are useful for inspecting embedded Office XML, manually recovering diagrams, and comparing add-in write-back content.

## Quick Install

1. Download `DrawioPpt-v1.0.8.zip` from the [v1.0.8 Release](https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.8).
2. Extract it to a local folder.
3. Close PowerPoint and Word.
4. Double-click `install.cmd`.
5. Open PowerPoint or Word and confirm that the `Draw.io` Ribbon appears.
6. Check the BuildId in the Ribbon info area or settings footer; the release installer prints the same BuildId.

To uninstall, close PowerPoint and Word, then run `uninstall.cmd`.

The release installer registers both the PowerPoint and Word add-ins. For repository-based development, build first and then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\register-office-addins.ps1
```

If you need to diagnose only one host, use `scripts\register-addin.ps1` or `scripts\register-word-addin.ps1` directly.

After installation, run this acceptance check to confirm both Word and PowerPoint are registered and loadable through Office COM:

```powershell
powershell -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Greensoft\DrawioPpt\scripts\verify-office-install.ps1"
```

The package root and `bin` folder contain `BuildInfo.txt`, and `PACKAGE.txt` records `BuildId` and `GitShortHash`. These fields help confirm whether Office loaded DLLs from the current release package.

## Documentation

- [User guide](./docs/user-guide.en.md)
- [Installation and local debugging guide](./docs/installation.en.md)
- [Word add-in design and usage](./docs/word-addin.en.md)
- [Architecture](./docs/architecture.en.md)
- [Regression checklist](./docs/regression-checklist.en.md)
- [v1.0.8 release notes](./docs/release-notes-v1.0.8.en.md)
- [Verified v1.0.8 E2E test report](./docs/e2e-test-report-v1.0.8.en.md)
- [v1.0.8 release evidence and log index](./docs/release-evidence-v1.0.8.en.md)

Development plans, historical release notes, and optimization backlog live under [docs](./docs/). The README is kept as the project overview and quick-start entry point.

## Technical Boundaries

- Target environment: Windows + Microsoft PowerPoint/Word Desktop.
- Office Web, WPS, and mobile Office are out of scope.
- Editing requires the add-in; users without it can still view diagrams, but cannot enter the Draw.io write-back workflow directly.
- The release package uses current-user COM registration and does not require administrator privileges.
