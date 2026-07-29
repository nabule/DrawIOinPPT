# DrawioPpt v1.0.8 Release Notes

[English](./release-notes-v1.0.8.en.md) | [中文](./release-notes-v1.0.8.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## Highlights

- The Word selection hot path no longer subscribes to `WindowSelectionChange`, and startup does not read selection, picture properties, `AlternativeText`, or `CustomXMLParts`. Ordinary selection, dragging, resizing, and repositioning perform no add-in metadata work.
- Re-edit, Refresh, Bind, and Clear Binding remain enabled. The user selects a picture and then clicks a command; only then does the add-in read the live selection and perform the operation. Selection alone never opens the editor.
- The normal Word display path renders a 1240px-wide, aspect-preserving PNG from the Draw.io source and explicitly requests draw.io `--border 0`. It converts the picture to a floating `Shape` with `wdWrapFront`. If draw.io Desktop export fails, Word generates a lightweight aspect-preserving PNG placeholder from a 1240px baseline and caps extreme portrait height at 1200px instead of falling back to complex SVG; source XML remains in document metadata and stays editable through an explicit command.
- PowerPoint continues to insert the original SVG directly. It uses proportional `contain` sizing within the slide bounds, centers the shape, and does not rewrite the `viewBox` or introduce internal blank padding to fill a box.
- Word Re-edit and Refresh retain picture width and recalculate height and position from the source ratio. Complete Draw.io XML remains in `Document.CustomXMLParts`; picture `AlternativeText` normally carries only a lightweight reference.
- If `CustomXMLParts.Upsert` fails, the picture fallback retains the full envelope. Legacy documents with a full envelope only in picture `AlternativeText` migrate it into document-level primary storage on first read.

## Compatibility Tradeoff

Office does not guarantee that copying a single Word picture into another document also copies the source document's `CustomXMLParts`. A normal lightweight `AlternativeText` reference contains no Draw.io XML, so the destination may be unable to recover the editing source. Copy the complete `.docx`, or retain the sidecar `.drawio` / Draw.io XML before cross-document delivery.

The Word PNG is a display artifact, not primary storage. The normal path requires working draw.io Desktop PNG export; export failure uses a lightweight PNG placeholder rather than complex SVG, without losing document-level source XML.

## Current Validation Results

- The `Release|x64` build passed with 0 warnings and 0 errors.
- The Word zero-passive-selection path, WebView2 user-data-folder contract, URL-host safety contract, and full-E2E cleanup safety contract passed.
- The latest `word-preview-image-provider-e2e.ps1` covers the normal 1240px aspect-preserving PNG, a 1240×620 lightweight PNG placeholder for its 2:1 failure sample, and immediate temporary-source, preview, and fallback-directory cleanup. The implementation still includes bounded retry, deferred, and startup cleanup for temporary XML, but this script does not claim coverage for those paths without fault injection.
- The latest `word-svg-aspect-ratio-e2e.ps1` passed in real Word: a 2:1 PNG remains 2:1 after create and replacement; the vertical case reports `VerticalRatio=0.25` and `VerticalContained=True`; a failed replacement reports `FailedReplacementPreservedOriginal=True`, proving the original is not deleted before the new picture succeeds.
- `powerpoint-svg-aspect-ratio-e2e.ps1` passed in real PowerPoint: the original SVG is inserted directly, centered with proportional `contain` sizing, and replacement retains source ratio without expanding the `viewBox`.
- Full functional Office E2E against the final package reported `12/12 PASS`, including `InstalledWordSvgAspectE2E`, `InstalledWordPreviewProviderE2E`, and `InstalledPowerPointSvgAspectE2E`.
- Before full E2E changes registration, it snapshots the Word and PowerPoint add-in registry trees; cleanup restores the exact prior state, including keys that were originally absent. The final zero-residual `WINWORD` / `POWERPNT` gate passed, with `FullE2ECleanupSucceeded=True`.
- Core, PowerPoint, and Word DLLs match across source Release output, the package, and the standard local installation. The Word DLL is `A6AB59AA…3492E8`. See the [release evidence](./release-evidence-v1.0.8.en.md) for the complete DLL hashes.

## Word Stress and Pointer Boundary

The final user-complex-source sample used a 1240×1753 PNG, floating `Front` layout, and a rich-document baseline with 150 body paragraphs, 19,456 body characters, eight pages, two tables, and three auxiliary pictures. In the latest run after blank-padding validation was corrected, normal/managed selection P95 values were `68.866/74.243 ms`, position P95 values were `321.844/325.764 ms`, and resize P95 values were `89.633/122.968 ms`; all `96/96` target selection events were confirmed, with none missing.

The latest run reports `ComparisonPassed=false`, `AbsoluteLatencyGatePassed=false`, and `OverallPassed=false`. Managed-picture round position P95 values were `303.892/337.339 ms`, while the normal control reported `328.256/318.748 ms`. An earlier confirmation run passed the relative gates and kept both managed rounds below 300ms, which shows environmental variation but is not grounds for relaxing the gate. All rounds had zero final failures and zero Word-unresponsive samples and must not be described as complete automated performance-acceptance passes.

The latest PNG is an opaque 24bpp raster with no alpha channel, so transparent-padding pixel inspection explicitly reports `BlankPaddingDetectionSupported=false` instead of falsely claiming a pixel-level pass. Current evidence against forced 1:1 sizing and export borders comes from the independent SVG source-ratio comparison (`0.000236` error) plus draw.io `--border 0`; transparent-PNG counterexamples are covered separately by the focused validation script.

`PointerInputCovered=false`. The test updates `Shape` properties on Word's visible STA UI thread and does not inject real mouse input. These release notes make no claim that mouse dragging passed. On 2026-07-29, the user chose to skip real-pointer drag, resize, reposition, and manual Re-edit acceptance; the status is “skipped/not accepted,” not “passed.”

## Release Status

v1.0.8 has been published as a GitHub Release: <https://github.com/nabule/DrawIOinPPT/releases/tag/v1.0.8>. The final package has been rebuilt, the standard local installation updated, and both `12/12` full E2E and the internal-link gate passed. ZIP SHA256 is not embedded in a document packaged into the same ZIP; the GitHub Release asset record and external delivery summary record it.

See the [v1.0.8 E2E test report](./e2e-test-report-v1.0.8.en.md), [v1.0.8 release evidence](./release-evidence-v1.0.8.en.md), and [Word UI-thread stress test report](./word-ui-thread-stress-report-v1.0.8.en.md) for details.
