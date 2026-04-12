# v1.0.0

[English](./release-notes-v1.0.0.en.md) | [中文](./release-notes-v1.0.0.md) | [English Home](../README.en.md) | [中文首页](../README.md)

This is the first stable `1.0` release of DrawioPpt. The goal of this release is to make the main workflow of keeping editable Draw.io diagrams maintained together with PowerPoint deliverable end-to-end, while also tightening installation, testing, logging, and release materials.

## Release Positioning

- A stable release for PowerPoint Desktop
- Continues the dual-mode architecture of `C# + COM Add-in + WebView2 + draw.io Desktop/URL`
- Focused on a deliverable workflow where source data stays with the PPT, desktop mode remains maintainable over time, and URL mode can write back on a real machine

## Highlights

### 1. Improved Ribbon Experience

- Ribbon regrouped into `Create / Current Shape / Workflow / Info`
- The main action button switches between `Edit` and `Re-edit` based on the current selection state
- Clearer selection-state, mode-state, and workflow prompts added
- Custom icons added for major buttons to reduce first-use friction

### 2. Automatic Sidecar Path Relocation

- When the PPT has not been saved yet, the sidecar still goes to the system temp directory
- After the PPT is saved, saved as, or moved, the add-in rebuilds the sidecar on the next edit using the current PPT path
- After copying the `pptx` to another machine, desktop mode prefers the new local path on that machine as long as the add-in and editor environment are ready

### 3. Persistence Model Further Tightened

- Draw.io source XML continues to be stored in `Presentation.CustomXMLParts`
- Shapes keep `diagramId / customXmlPartId` references
- Legacy `AlternativeText` fallback data remains supported and is migrated automatically during reads
- SVG continues to receive draw.io `content` metadata for recovery and troubleshooting

### 4. More Complete Release Materials

- The release package still includes install scripts, uninstall scripts, and user-facing documentation
- A new release-evidence document summarizes build, packaging, E2E, and log artifacts
- The release package carries version-matched test logs for delivery and traceability

## Compatibility Notes

- Supports both `Desktop` and `Url` editing modes
- Still requires Windows + PowerPoint Desktop x64 + .NET Framework 4.8 + WebView2 Runtime
- URL mode still depends on the target address supporting the draw.io embed protocol
- Desktop mode still works best when `draw.io Desktop` or `diagrams.net Desktop` is installed locally

## Upgrade Recommendations

- If you are upgrading from `v0.6.1-stable`, reinstalling the `v1.0.0` release package is recommended
- If older documents point sidecar files to a temp directory or an old machine path, `v1.0.0` will prefer switching back to the current PPT path on the next edit
- After upgrading, it is recommended to run at least one full check of:
  - Create a new diagram
  - Save the PPT
  - Close and reopen the PPT
  - Edit the same diagram again

## Companion Documents

- [Installation and local debugging guide](./installation.en.md)
- [User guide](./user-guide.en.md)
- [v1.0.0 E2E test report](./e2e-test-report-v1.0.0.en.md)
- [v1.0.0 release evidence and log index](./release-evidence-v1.0.0.en.md)
