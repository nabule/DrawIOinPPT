# DrawioPpt v1.0.8 Release Notes

[English](./release-notes-v1.0.8.en.md) | [中文](./release-notes-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Highlights

- The Word selection hot path no longer subscribes to `WindowSelectionChange`, and startup does not read selection, picture properties, `AlternativeText`, or `CustomXMLParts`. Ordinary selection, dragging, resizing, and repositioning perform no add-in metadata work.
- Re-edit, Refresh, Bind, and Clear Binding remain enabled. The user selects a picture and then clicks a command; only then does the add-in read the live selection and perform the operation. Selection alone never opens the editor.
- The normal Word display path renders a 620px-wide, border-free, aspect-preserving PNG from the Draw.io source. It converts the picture to a floating `Shape` with `wdWrapFront`. If draw.io Desktop export fails, Word generates a lightweight aspect-preserving 620×310 PNG placeholder instead of falling back to complex SVG; source XML remains in document metadata and stays editable through an explicit command.
- PowerPoint continues to insert the original SVG directly. It uses proportional `contain` sizing within the slide bounds, centers the shape, and does not rewrite the `viewBox` or introduce internal blank padding to fill a box.
- Word Re-edit and Refresh retain picture width and recalculate height and position from the source ratio. Complete Draw.io XML remains in `Document.CustomXMLParts`; picture `AlternativeText` normally carries only a lightweight reference.
- If `CustomXMLParts.Upsert` fails, the picture fallback retains the full envelope. Legacy documents with a full envelope only in picture `AlternativeText` migrate it into document-level primary storage on first read.

## Compatibility Tradeoff

Office does not guarantee that copying a single Word picture into another document also copies the source document's `CustomXMLParts`. A normal lightweight `AlternativeText` reference contains no Draw.io XML, so the destination may be unable to recover the editing source. Copy the complete `.docx`, or retain the sidecar `.drawio` / Draw.io XML before cross-document delivery.

The Word PNG is a display artifact, not primary storage. The normal path requires working draw.io Desktop PNG export; export failure uses a lightweight PNG placeholder rather than complex SVG, without losing document-level source XML.

## Current Validation Results

- The `Release|x64` build passed with 0 warnings and 0 errors.
- The Word zero-passive-selection path, WebView2 user-data-folder contract, URL-host safety contract, and full-E2E cleanup safety contract passed.
- The latest `word-preview-image-provider-e2e.ps1` covers the normal 620px aspect-preserving PNG, the 620×310 aspect-preserving lightweight PNG placeholder on failure, and temporary-XML retry/deferred/startup cleanup.
- The latest `word-svg-aspect-ratio-e2e.ps1` passed in real Word: a 2:1 PNG remains 2:1 after create and replacement; the vertical case reports `VerticalRatio=0.25` and `VerticalContained=True`; a failed replacement reports `FailedReplacementPreservedOriginal=True`, proving the original is not deleted before the new picture succeeds.
- `powerpoint-svg-aspect-ratio-e2e.ps1` passed in real PowerPoint: the original SVG is inserted directly, centered with proportional `contain` sizing, and replacement retains source ratio without expanding the `viewBox`.
- Full functional Office E2E against the final package reported `12/12 PASS`, including `InstalledWordSvgAspectE2E`, `InstalledWordPreviewProviderE2E`, and `InstalledPowerPointSvgAspectE2E`.
- Before full E2E changes registration, it snapshots the Word and PowerPoint add-in registry trees; cleanup restores the exact prior state, including keys that were originally absent. The final zero-residual `WINWORD` / `POWERPNT` gate passed, with `FullE2ECleanupSucceeded=True`.
- Core, PowerPoint, and Word DLLs match across source Release output, the package, and the standard local installation. The Word DLL is `6E6DEAF2…B127D`. See the [release evidence](./release-evidence-v1.0.8.en.md) for the complete DLL hashes.

## Word Stress and Pointer Boundary

The final user-complex-source sample used a 620×876 PNG, floating `Front` layout, and a rich-document baseline. Normal/managed position P95 values were `524.962/488.881 ms`, with the managed picture `36.081 ms` lower; all `98/98` target selection events were confirmed, with none missing.

The absolute position-P95 gate was `300 ms`, and both the normal and managed pictures failed it. A `17×24 px` solid-color PNG baseline also failed the same gate. The machine result is `ComparisonPassed=false`, `AbsoluteLatencyGatePassed=false`, and `OverallPassed=false`; it must not be described as an automated performance-acceptance pass.

`PointerInputCovered=false`. Windows denied `GetCursorPos`, and the `SetIsBorderRequired` interface is unsupported, so automation did not cover real mouse input. These release notes make no claim that mouse dragging passed. Real-pointer drag, resize, reposition, and a manual Re-edit click remain separate pending acceptance items.

## Release Status

v1.0.8 remains pending public release, but the final package has been rebuilt, the standard local installation updated, and both `12/12` full E2E and the internal-link gate passed. ZIP SHA256 is not embedded in a document packaged into the same ZIP; the external delivery summary records it.

See the [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md), [v1.0.8 release evidence](./release-evidence-v1.0.8.en.md), and [Word UI-thread stress test report](./word-ui-thread-stress-report-v1.0.8.en.md) for details.
