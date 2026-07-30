# v1.0.3

[English](./release-notes-v1.0.3.en.md) | [中文](./release-notes-v1.0.3.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.3` is a release for the Word add-in delivery and the recommended desktop editor configuration. It documents the Word COM Add-in, the in-document Draw.io XML storage model, and the draw.io Desktop build with XML clipboard tools.

## Problems Solved First

- Users needed the same Draw.io maintenance workflow in Word, not only in PowerPoint.
- When delivering `.pptx` / `.docx` files, users needed to know whether the editable source XML actually stayed inside the Office file.
- Troubleshooting or recovering diagrams was difficult without an easy way to paste embedded XML back into draw.io Desktop.
- Release packages should not include debug symbols or historical logs that may expose local paths.

## Highlights

- Adds Word Desktop COM Add-in delivery documentation for creating, re-editing, refreshing, binding, and clearing Draw.io pictures.
- Clarifies that Draw.io source XML is embedded in document-level `CustomXMLParts` inside `.pptx` / `.docx`; the sidecar `.drawio` file is an editing cache, manual backup, and auto-refresh helper.
- Adds [draw.io Desktop v30.0.4 XML tools build](https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1) setup guidance to the README, user guide, installation guide, and Word-specific guide.
- Release packages no longer include PDB files or historical logs by default, reducing local path/debug information in distributed packages.

## Recommended draw.io Desktop Version

Recommended download:

```text
draw.io-30.0.4-xml-tools.1-windows-x64-unpacked.zip
```

Release page:

```text
https://github.com/nabule/drawio-desktop/releases/tag/v30.0.4-xml-tools.1
```

SHA256:

```text
B09116FB0D6140E39CFDA0568957897BD697CB5B7DAE257F7B1438E9B1A6DB9D
```

This is an unsigned Windows x64 unpacked build. After extraction, configure `Desktop Path` to:

```text
win-unpacked\draw.io.exe
```

Do not copy only the single exe.

## XML Tool Buttons

This draw.io Desktop build is based on official `v30.0.4` and adds two colored XML icon buttons to the second toolbar row:

- Paste draw.io XML source from the clipboard and render it as a diagram.
- Copy the current canvas as draw.io XML source.

This helps DrawioPpt desktop-mode diagnostics: you can put the Draw.io XML embedded in an Office document back into draw.io Desktop, or copy the current canvas as XML for comparison, recovery, or testing.

## Validation

- Release build passed.
- Release package generation passed.
- Temporary package installation and uninstall passed.
- Package content inspection confirmed no PDB, log, or screenshot assets.

Details:

- [v1.0.3 E2E test report](./e2e-test-report-v1.0.3.en.md)
- [v1.0.3 release evidence and log index](./release-evidence-v1.0.3.en.md)
