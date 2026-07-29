# Regression Checklist

[English](./regression-checklist.en.md) | [中文](./regression-checklist.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## v1.0.8 Execution Status

| Status | Item |
| --- | --- |
| PASS | Full Office E2E against the temporary release-package installation passed 12/12; this result covers only the isolated temporary installation. |
| PASS | Standard local installation at `%LOCALAPPDATA%\Greensoft\DrawioPpt`, DLL hashes, Office registration, and real COM loading. |
| Absolute gate failed | In the final 620px PNG / `Front` rich-document stress sample, normal/managed position P95 values were 524.962/488.881 ms, with the managed picture 36.081 ms lower and 98/98 target selection events confirmed. Both failed the absolute 300 ms gate, as did the 17×24 solid-color PNG baseline. |
| Pending manual confirmation | `PointerInputCovered=false`. Windows denied `GetCursorPos`, and the `SetIsBorderRequired` interface is unsupported, so automation cannot substitute for real-pointer drag, resize, reposition, and command-click acceptance. |

## 1. Build and Registration

- `scripts/dev-check.ps1` passes
- `scripts/build.ps1` passes
- `scripts/register-office-addins.ps1` registers both PowerPoint and Word successfully
- `Greensoft.DrawioPptAddIn` loads correctly in PowerPoint

## 2. Ribbon and Selection State

- The Ribbon displays correctly
- Ribbon groups are `Create / Current Shape / Workflow / Info`
- Word status remains `Select a picture, then click an action / Identified when a button is clicked`; selection changes do not invalidate the Ribbon
- Word Re-edit, Refresh, Bind, and Clear Binding remain enabled and validate the live selection on click
- Word does not show Auto Open; PowerPoint `AutoOpenOnSelection` behavior is unchanged
- Button icons render correctly without blank placeholders

## 3. Desktop Mode

- Local `draw.io` path can be auto-detected
- A new diagram can be created
- The desktop editor can be opened
- When a new diagram is created in an unsaved blank PPT, the sidecar is written to an absolute temp path instead of generating a relative `drawio\*.drawio`
- After the PPT is saved or saved as, the sidecar switches automatically to the current PPT directory on the next edit
- After copying the PPT to another machine, editing again rebuilds the sidecar based on the PPT path on that machine
- Saving the `.drawio` file refreshes the SVG in PPT automatically
- The manual `Refresh` button works

## 4. URL Mode

- `Test URL` passes in the settings window
- Existing diagrams can be opened
- The `init -> load -> save -> export` flow completes
- SVG and XML can be written back to PowerPoint
- After creating a new diagram, saving the PPT, closing it, and reopening it, the same diagram can still be edited and written back again
- The log records the key URL-mode events
- `scripts\word-url-addin-host-e2e.ps1` passes inside real `WINWORD.EXE`: the add-in connects, a managed picture is selected, Re-edit is invoked explicitly, write-back completes, `ExplicitEditCommandInvoked=True`, `%LOCALAPPDATA%\Greensoft\DrawioPpt\WebView2` exists, and the new log segment has no `E_ACCESSDENIED`.

## 5. Storage and Migration

- New diagrams write data into `Shape.Tags`
- New diagrams write data into `Presentation.CustomXMLParts`
- After reopening the PPT, source XML can be restored from `CustomXMLParts` and editing can continue
- Legacy shapes that only contain `AlternativeText` data are migrated automatically
- After clearing shape binding, orphaned `CustomXMLPart` entries with no references are removed
- Generated SVG includes draw.io `content` metadata

## 6. Display-object Replacement Fidelity

- Position is preserved after replacement
- Size is preserved after replacement
- Rotation is preserved after replacement
- Shape name is preserved after replacement
- Hyperlinks are preserved after replacement
- `scripts\word-preview-image-provider-e2e.ps1` covers the normal 620px aspect-preserving PNG, the 620×310 aspect-preserving lightweight PNG placeholder on export failure, and temporary-XML retry/deferred/startup cleanup. It must not fall back to complex SVG.
- `scripts\word-svg-aspect-ratio-e2e.ps1` passes in real Word: a 2:1 PNG remains 2:1 after insertion and replacement, both pictures are floating `wdWrapFront` shapes, `VerticalRatio=0.25`, `VerticalContained=True`, and `FailedReplacementPreservedOriginal=True` proves create-new-before-delete-old replacement.
- `scripts\powerpoint-svg-aspect-ratio-e2e.ps1` passes in real PowerPoint: the original SVG is inserted directly, centered with proportional `contain` sizing, and replacement retains source ratio without expanding the `viewBox` or adding internal blank space.

## 7. Exceptions and Logging

- Missing desktop path produces a clear prompt
- URL initialization timeout produces a clear prompt
- URL export timeout produces a clear prompt
- Host entry-point exceptions are logged and also shown through an error dialog
- Log files continue to append correctly

## 8. Registration and Uninstall

- The uninstall script removes registration entries correctly
- After uninstall, PowerPoint and Word no longer load the add-ins

## 9. Word Selection and Picture Operations

- `scripts\word-selection-sync-test.ps1` passes: `SelectionMonitor` contains no `WindowSelectionChange` / `SelectionChanged`, and `AddInHost` contains no selection timer, passive-selection handler, or startup selection read.
- `scripts\word-selection-event-e2e.ps1` passes against real Word COM: the release DLL reports `NoPassiveSelectionMetadataPath=True`, while explicit reads still identify a managed floating `Shape` and a normal picture.
- Close Word before running the Word E2E and do not run it alongside another Word automation test. The script restores add-in settings and cleans up only automation Word processes created by this test.
- Ordinary selection, movement, resizing, and repositioning must not read picture `AlternativeText`, `CustomXMLParts`, or invalidate the Ribbon. Select the picture first, then click Re-edit, Refresh, Bind, or Clear Binding; only the command reads the live selection.

## 10. Word Complex Metadata and Drag Performance

Automated checks:

- `scripts\word-complex-metadata-e2e.ps1` uses a large Draw.io XML payload to verify that the full body is stored in `Document.CustomXMLParts`, while picture `AlternativeText` contains no `DrawioXml` and is no longer than 2048 characters.
- After saving, closing, and reopening the `.docx`, both the document-level full envelope and the lightweight picture reference still exist with consistent reference fields.
- When `CustomXMLParts.Upsert` failure is simulated, picture `AlternativeText` retains the full envelope, which remains readable after save, close, and reopen.
- A legacy picture with its full envelope only in `AlternativeText` migrates into `CustomXMLParts` on first read; a second read creates no new part and keeps the related XML part ID stable.
- `scripts\word-url-e2e.ps1` strictly verifies URL-mode create, reopen, edit, and final write-back through `Document.CustomXMLParts`; the lightweight `AlternativeText` must not stand in for the full primary store.
- `scripts\word-url-addin-host-e2e.ps1` strictly verifies actual-host write-back through the document-level primary store, with `WordAddInConnect=True` and `ActualWordUrlEditorSaved=True`.
- Each Word automation script checks that no residual `WINWORD.EXE` remains; a residual process fails the check.
- The installed visible-Word UI-thread stress comparison rendered the user-supplied complex source as a 620×876 PNG; normal and managed pictures were proportional floating `Front` pictures. Normal/managed position P95 values were 524.962/488.881 ms, with the managed picture 36.081 ms lower and all 98/98 target-Shape selection events confirmed. The absolute 300 ms gate failed, and a 17×24 solid-color PNG baseline failed it as well; `ComparisonPassed=false`, `AbsoluteLatencyGatePassed=false`, and `OverallPassed=false`. Installed-build runtime reflection independently confirms `NoPassiveSelectionMetadataPath=True`, but that source/runtime fact is not a performance or mouse-acceptance pass. See the [v1.0.8 Word UI-thread stress test report](./word-ui-thread-stress-report-v1.0.8.en.md).

Pre-release manual acceptance:

| Status | Check |
| --- | --- |
| Not covered by automation; pointer confirmation pending | `PointerInputCovered=false`; Windows denied or did not support the pointer interfaces, and the absolute 300 ms gate failed. The user must still drag, resize, and reposition each object with a real pointer for 10 continuous seconds each to confirm visual tracking. |
| Automated explicit command passed; manual button confirmation pending | Save and reopen preserved position, size, and all 26,106 source-XML characters, and the real URL-host E2E completed write-back through select-then-explicit-Re-edit. This is not user-click acceptance; one local manual Re-edit click is still required. |
