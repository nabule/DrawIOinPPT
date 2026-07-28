# v1.0.8

[English](./release-notes-v1.0.8.en.md) | [中文](./release-notes-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

`v1.0.8` optimizes the picture-metadata read path for complex Draw.io diagrams in Word. The complete Draw.io XML remains embedded in the `.docx`, but is no longer duplicated in normal picture `AlternativeText`. This removes large-XML deserialization from `WindowSelectionChange` synchronization and improves responsiveness while selecting, dragging, or resizing complex diagrams.

## Highlights

- On a normal Word write, `Document.CustomXMLParts` remains the in-document primary store for the full envelope and compressed `DrawioXml`.
- `InlineShape/Shape.AlternativeText` is now a lightweight envelope without `DrawioXml`; it retains only `formatVersion`, `diagramId`, diagram name, editor mode, editor target, sidecar path, and update time.
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
- `word-selection-event-e2e.ps1` passes: real Word selection recognition and the no-polling contract satisfy their assertions.
- `word-url-e2e.ps1` passes: the ordinary URL create, reopen, and edit flow verifies full data through `Document.CustomXMLParts`.
- `word-url-addin-host-e2e-safety-test.ps1` passes: actual-host script constraints for registration, settings, process ownership, and cleanup satisfy their assertions.
- `word-url-addin-host-e2e.ps1` passes in real `WINWORD.EXE`, with `WordAddInConnect=True` and `ActualWordUrlEditorSaved=True`.
- The full Office E2E against the temporary v1.0.8 installation passes 9/9. It adds `InstalledWordComplexMetadataE2E` and covers packaging, installation, Office registration/load, URL smoke, the real Word URL host, PowerPoint URL write-back, and desktop export.
- The v1.0.8 package has been generated with `scripts\word-complex-metadata-e2e.ps1` and the Chinese and English reports and evidence indexes. See the [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md) and [v1.0.8 release evidence](./release-evidence-v1.0.8.en.md).
- Full E2E identifies and cleans up only its test-owned Word process by PID and start time. Temporary-package uninstall, repository add-in re-registration, or settings-restoration failures propagate to the final failure state; the cleanup safety contract passes.
- The package explicitly includes both architecture documents, regression checklists, optimization backlogs, and historical v1.0.5 E2E reports. Its pre-compression Markdown-link check reports zero missing targets.

v1.0.8 remains pending public release. Final installation into the standard local directory and the manual 10-second drag, resize, reposition, and reopen/edit comparison remain later acceptance steps and are not claimed here.
