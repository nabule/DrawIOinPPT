# v1.0.6

[English](./release-notes-v1.0.6.en.md) | [中文](./release-notes-v1.0.6.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.6` improves responsiveness while Draw.io pictures in Word are moved, resized, or repositioned. SVG, Draw.io XML, picture format, the `AlternativeText`/`CustomXMLParts` storage model, and floating-picture layout behavior are unchanged.

## Highlights

- Removes the Word host's 500 ms `System.Windows.Forms.Timer`, so picture metadata is no longer parsed periodically.
- Retains `WindowSelectionChange` and one selection read at add-in startup for Ribbon state and enabled auto-open behavior.
- Edit, Refresh, Bind, and Clear Binding always read the live Word selection on click, so a manual command still targets the selected picture when Word has not immediately refreshed selection display.
- Adds a no-polling source contract, a real Word COM floating-picture selection test, and a GitHub Actions contract check.
- Fixes the Word URL E2E: the mock server is now self-hosted by the test process with an HTTP reachability check; the test isolates the registered add-in and restores its original load setting to avoid a double host.
- Includes `word-selection-event-e2e.ps1` in the release package for real Word COM validation of the release DLL.

## Usage Note

Word still performs floating-picture layout itself; this release removes only the add-in's half-second background selection access. If Ribbon state does not refresh immediately, select the picture and invoke Re-edit, Refresh, Bind, or Clear Binding; the command reads the live selection.

Close Word and do not run another Word automation task before running `word-url-e2e.ps1` or `word-selection-event-e2e.ps1`. See the [v1.0.6 E2E test report](./e2e-test-report-v1.0.6.en.md) for verification details.
