# v1.0.8

[English](./release-notes-v1.0.8.en.md) | [中文](./release-notes-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.8` changes Word complex-picture interaction to prioritize zero add-in drag overhead. Complete Draw.io XML remains embedded in the `.docx`, while normal picture `AlternativeText` carries only a lightweight reference. Word also removes its `WindowSelectionChange` subscription and startup selection read entirely. Selection, dragging, resizing, and repositioning perform no add-in metadata detection or Ribbon refresh; users select the picture first and then click Re-edit or another command.

## Highlights

- On a normal Word write, `Document.CustomXMLParts` remains the in-document primary store for the full envelope and compressed `DrawioXml`.
- `InlineShape/Shape.AlternativeText` is now a lightweight envelope without `DrawioXml`; it retains only `formatVersion`, `diagramId`, diagram name, editor mode, editor target, sidecar path, and update time.
- Word no longer monitors ordinary selection changes and does not read selection, picture metadata, or scan document orphans at startup.
- Re-edit, Refresh, Bind, and Clear Binding remain enabled and read and validate the live selection only when clicked. The Ribbon displays fixed select-then-click guidance.
- Word removes Auto Open; PowerPoint selection events and `AutoOpenOnSelection` are unchanged. Double-clicking a Word picture remains an explicit editing action.
- If `CustomXMLParts.Upsert` fails, picture `AlternativeText` retains the full envelope so source data is not lost during the save.
- A legacy document that stores its full envelope only in picture `AlternativeText` is migrated into `CustomXMLParts` on first read. After successful migration the picture becomes a lightweight reference, and repeated reads do not create new XML parts.
- The ordinary Word URL regression and actual add-in host regression now strictly verify the document-level primary store, preventing a lightweight reference from being mistaken for complete write-back.
- The PowerPoint `Presentation.CustomXMLParts + Shape.Tags + Shape.AlternativeText` storage strategy is unchanged by this Word-specific optimization.

## Compatibility Tradeoff

Office does not guarantee that copying one Word picture into another document also copies the source document's `CustomXMLParts`. Because normal lightweight `AlternativeText` does not contain Draw.io XML, the destination document may not be able to recover the editing source.

For cross-document delivery, copy the complete `.docx`, or save/retain the sidecar `.drawio` or Draw.io XML first. Only legacy pictures and primary-store failure fallbacks may still carry the full payload in `AlternativeText`; it is not a guaranteed recovery source for every picture.

## Validation Completed So Far

- The Release x64 solution build passes with `0` warnings and `0` errors.
- `word-complex-metadata-e2e.ps1` passes: large-XML primary storage, lightweight reference, save-close-reopen persistence, failure fallback, legacy migration, stable part ID on the second read, and no residual Word process all satisfy their assertions.
- `word-selection-event-e2e.ps1` passes: the release DLL reports `NoPassiveSelectionMetadataPath=True`, while explicit reads still identify managed and normal floating pictures.
- `word-url-e2e.ps1` passes: the ordinary URL create, reopen, and edit flow verifies full data through `Document.CustomXMLParts`.
- `word-url-addin-host-e2e-safety-test.ps1` passes: actual-host script constraints for registration, settings, process ownership, and cleanup satisfy their assertions.
- `word-url-addin-host-e2e.ps1` passes in real `WINWORD.EXE` through the select-then-explicit-Re-edit flow, with `WordAddInConnect=True`, `ExplicitEditAutomationAvailable=True`, `ExplicitEditCommandInvoked=True`, and `ActualWordUrlEditorSaved=True`.
- The full Office E2E against the temporary v1.0.8 installation passes 9/9. It adds `InstalledWordComplexMetadataE2E` and covers packaging, installation, Office registration/load, URL smoke, the real Word URL host, PowerPoint URL write-back, and desktop export.
- The v1.0.8 package has been generated with `scripts\word-complex-metadata-e2e.ps1` and the Chinese and English reports and evidence indexes. See the [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md) and [v1.0.8 release evidence](./release-evidence-v1.0.8.en.md).
- Full E2E identifies and cleans up only its test-owned Word process by PID and start time. Build, package, install, compilation, temporary-package uninstall, repository add-in re-registration, and settings-restoration failures all enter the common catch/finally path, FatalError report row, and trailing `exit 1`; the cleanup and unified-exit safety contract passes.
- The package explicitly includes both architecture documents, regression checklists, optimization backlogs, and historical v1.0.5 E2E reports. Its pre-compression Markdown-link check reports zero missing targets.
- v1.0.8 is installed in the standard per-user directory `%LOCALAPPDATA%\Greensoft\DrawioPpt`. `PACKAGE.txt` reports `v1.0.8 / Release / x64`; the three DLL SHA256 values match the release package, the Word and PowerPoint `CodeBase` values point to that standard directory, and add-in loading passed.
- A real Word UI-thread stress comparison against the installed build used 120 SVG elements and 254,606 characters of Draw.io XML. Two crossed-order rounds used a four-second warmup and 12-second measurement per picture per round, recording selection, position, and size latency separately. Combined normal/managed P95 values were 7.501/5.833 ms, 388.910/368.568 ms, and 112.280/103.247 ms, with managed/normal ratios of 0.778, 0.948, and 0.920. The comparison, absolute-latency, minimum-sample, final-failure, and responsiveness gates passed; the add-in was connected in the measured instance, 135 target-Shape events were classified with zero missing, and installed-build runtime reflection reported `NoPassiveSelectionMetadataPath=True`. This sample supports only that no additional difference above the threshold was observed and that absolute latency remained inside this test's gates; it does not establish the cause of the absolute latency.
- After save and reopen, the managed picture retained its position and size, its lightweight `AlternativeText` remained 382 characters, and `Document.CustomXMLParts` restored all 254,606 characters of source XML. The real URL-host regression also verifies `ActualWordUrlEditorSaved=True`. See the [Word UI-thread stress acceptance report](./word-ui-thread-stress-report-v1.0.8.en.md) for the full command and path-normalized evidence snapshot.

v1.0.8 remains pending public release. Standard local installation and the automated Word UI-thread stress comparison are complete. The current Codex interactive desktop is denied cursor access by the operating system (Win32 error 5), so it cannot substitute for the user's real pointer drag. A manual 10-second pointer drag, resize, and reposition on each normal and managed picture, followed by a Re-edit button click, therefore remain pending, and this document does not represent the automated stress result as manual mouse acceptance.
