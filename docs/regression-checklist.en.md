# Regression Checklist

[English](./regression-checklist.en.md) | [中文](./regression-checklist.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Build and Registration

- `scripts/dev-check.ps1` passes
- `scripts/build.ps1` passes
- `scripts/register-office-addins.ps1` registers both PowerPoint and Word successfully
- `Greensoft.DrawioPptAddIn` loads correctly in PowerPoint

## 2. Ribbon and Selection State

- The Ribbon displays correctly
- Ribbon groups are `Create / Current Shape / Workflow / Info`
- When no shape is selected, status shows `No selection`
- When a normal shape is selected, status shows that it is a normal shape and can be bound
- When a bound shape is selected, status shows that it is recognized as a Draw.io shape
- The main action button changes its label to `Re-edit` for a bound shape
- The `Auto Open` toggle reads and writes `AutoOpenOnSelection` correctly
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

## 5. Storage and Migration

- New diagrams write data into `Shape.Tags`
- New diagrams write data into `Presentation.CustomXMLParts`
- After reopening the PPT, source XML can be restored from `CustomXMLParts` and editing can continue
- Legacy shapes that only contain `AlternativeText` data are migrated automatically
- After clearing shape binding, orphaned `CustomXMLPart` entries with no references are removed
- Generated SVG includes draw.io `content` metadata

## 6. SVG Replacement Fidelity

- Position is preserved after replacement
- Size is preserved after replacement
- Rotation is preserved after replacement
- Shape name is preserved after replacement
- Hyperlinks are preserved after replacement

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

- `scripts\word-selection-sync-test.ps1` passes; `AddInHost` must not reintroduce a selection `Timer` or Tick handler.
- `scripts\word-selection-event-e2e.ps1` passes against real Word COM, validating managed floating-`Shape` recognition, the live-selection prerequisite for binding a normal picture, and the no-polling contract of the release DLL.
- Close Word before running the Word E2E and do not run it alongside another Word automation test. The script restores add-in settings and cleans up only automation Word processes created by this test.
- After a floating picture is moved or resized, Word events update Ribbon state. If Word does not update the display immediately, select the picture and invoke Re-edit, Refresh, Bind, or Clear Binding; the command reads the live selection.
