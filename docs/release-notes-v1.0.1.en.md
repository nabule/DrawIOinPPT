# v1.0.1

[English](./release-notes-v1.0.1.en.md) | [中文](./release-notes-v1.0.1.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.1` is a maintenance release focused on release-pipeline reliability and SVG rendering fidelity. The main goal is to fix the issue where diagrams edited in draw.io could look blurry in PowerPoint even though they were expected to stay vector-sharp.

## Problems Solved First

- Text in diagrams could look blurry in PowerPoint after editing in draw.io and writing back, which was not acceptable for formal presentations.
- Local build, packaging, installation, and E2E scripts could fail on Windows machines where a plain `powershell` command was unavailable.
- URL-mode smoke tests depended too much on external environment details, making release verification less repeatable.

## Highlights

### 1. Fix for blurry text after Draw.io write-back

- Some SVG exported by draw.io Desktop carries text as `foreignObject + data:image/png` bitmap fallback
- The add-in now converts that fallback into real SVG `<text>/<tspan>` content during SVG sanitization whenever it can
- After the diagram is written back into PowerPoint, the SVG stored inside the PPT package no longer keeps that bitmap text fallback, so zooming is much more consistent

### 2. SVG insertion stays unchanged at the workflow level

- The add-in still inserts and replaces diagram visuals as SVG
- The new compatibility logic lives in the SVG sanitization layer, so the existing metadata, sidecar, URL-mode, and desktop-mode workflows stay intact
- No new release-layer change was introduced to the existing shape geometry, sizing, rotation, or basic action-transfer behavior

### 3. Local release scripts now run reliably on the current machine

- `build.ps1`, `package-release.ps1`, `install-release.ps1`, `uninstall-release.ps1`
- `url-editor-smoke.ps1`, `full-e2e-test.ps1`
- Internal secondary invocations in these scripts now use `powershell.exe` explicitly so local build/install flows do not fail when the plain `powershell` alias is unavailable
- The local mock editor host used by URL smoke and full E2E now runs on an in-script `HttpListener`, so the verification flow no longer depends on an external Python installation

## Upgrade Notes

- If an older add-in build is already registered on your machine, re-running registration or installation is enough to move to `v1.0.1`
- If you previously saw blurry text after editing a diagram in draw.io and writing it back into PowerPoint, re-editing and saving that shape again with `v1.0.1` is recommended
- If you rely on the local packaging or full E2E scripts, use the current repository version of those scripts together with this release

## Companion Documents

- [v1.0.1 E2E test report](./e2e-test-report-v1.0.1.en.md)
- [v1.0.1 release evidence and log index](./release-evidence-v1.0.1.en.md)
- [Installation and local debugging guide](./installation.en.md)
- [User guide](./user-guide.en.md)
