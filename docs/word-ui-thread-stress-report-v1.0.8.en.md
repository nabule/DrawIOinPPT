# v1.0.8 Word UI-thread Stress Test Report

[English](./word-ui-thread-stress-report-v1.0.8.en.md) | [中文](./word-ui-thread-stress-report-v1.0.8.md) | [Sanitized machine-readable evidence](./evidence/word-ui-thread-stress-v1.0.8.json)

## Purpose and Boundary

This test runs in real Word against the standard local installation. It renders the same user-supplied complex Draw.io source as an aspect-preserving `620 px`-wide PNG, then creates a normal picture and a managed picture with matching visuals, dimensions, anchor semantics, and wrapping. Both are floating `Shape` objects with `Front` wrapping. The script applies the same selection, position, and size property updates on Word's visible STA UI thread to compare observable managed-metadata overhead.

The standard local Word DLL measured in this run has SHA256 `6E6DEAF2D5DFFDF9363FD28787E40E7AE681290CB669B4A64D82ED2D364B127D`, matching source Release output and the final package.

The test updates `Shape` properties through Word COM; it does not inject a real mouse pointer. `PointerInputCovered=false`. A separate attempt to use the Windows pointer interface was denied by the operating system with Win32 error 5. This report is therefore not mouse-drag acceptance and never represents the automated result as “real mouse passed.”

The Word add-in does not subscribe to `WindowSelectionChange`. An independent counter verifies the target-picture events raised by Word. During normal use, the user selects a picture and then clicks Re-edit or another command; only that command reads the live selection and starts editing. Selection, dragging, or resizing alone never opens the editor.

## Reproducible Command

Close all Word windows first. `<USER_DRAWIO_SOURCE>` denotes a caller-supplied local file; the report and evidence do not retain its absolute attachment path:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 `
  -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" `
  -DrawioSourcePath "<USER_DRAWIO_SOURCE>" `
  -PictureWrapMode Front `
  -PictureRenderFormat Png `
  -PreviewPixelWidth 620 `
  -DurationSeconds 10 `
  -MaximumP95Ratio 1.25 `
  -MaximumPerRoundP95Ratio 1.25 `
  -MaximumP95DeltaMs 50 `
  -MaximumSelectionP95DeltaMs 50 `
  -MaximumAbsoluteSelectionP95Ms 300 `
  -MaximumAbsoluteP95Ms 300 `
  -MaximumAbsoluteResizeP95Ms 300 `
  -MinimumOperationsPerRound 10 `
  -WarmupSeconds 4 `
  -ResizeOperationsPerRound 12 `
  -SelectionSettleMilliseconds 75 `
  -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8.json"
```

The script runs two crossed-order rounds—normal then managed, followed by managed then normal—in a document containing 150 body paragraphs, 19,456 body characters, eight pages, two tables, and three auxiliary pictures. The source contains 66 `mxCell` elements. The PNG is `620×876 px`; source and displayed aspect ratios are approximately `0.707763`, with no border pixels. The normal and managed pictures have the same size, anchor semantics, and `Front` wrapping.

## Results

Execution date: 2026-07-29.

| Metric | Normal picture | Managed picture |
| --- | ---: | ---: |
| Combined duration over two rounds | 20.895 s | 20.624 s |
| Selection operations | 49 | 49 |
| Selection P95 | 37.136 ms | 49.044 ms |
| Position updates | 25 | 25 |
| Mean position-update latency | 363.771 ms | 344.120 ms |
| Position-update P95 | 524.962 ms | 488.881 ms |
| Final position-update failures | 0 | 0 |
| Size updates | 24 | 24 |
| Size-update P95 | 83.134 ms | 71.072 ms |
| Final size-update failures | 0 | 0 |
| Word unresponsive samples | 0 | 0 |

- The combined managed/normal position-P95 ratio is `0.931`, with a delta of `-36.081 ms`; neither managed position updates nor managed size updates were slower than the normal picture.
- Round 1 normal/managed position P95 values are `540.403/499.107 ms`, a `-41.296 ms` delta. Round 2 values are `522.370/488.881 ms`, a `-33.489 ms` delta.
- The independent counter confirmed all `98` expected target-picture selection events, with zero missing: `98/98` and `SelectionChangeEventsPassed=true`.
- Save and reopen preserved aspect ratio, position, size, and all 26,106 source-XML characters in `Document.CustomXMLParts`; `WordAddInConnect=true`, and cleanup left no residual `WINWORD`.
- The absolute position-P95 gate is `300 ms`, and both the normal and managed pictures failed it. A separate `17×24 px` solid-color PNG baseline also failed the same absolute gate. The absolute latency therefore cannot be attributed to Draw.io complexity or managed metadata, and the automated absolute-performance gate did not pass.
- The machine result is `ComparisonPassed=false`, `AbsoluteLatencyGatePassed=false`, and `OverallPassed=false`. These stress gates are distinct from the functional full Office E2E result of `12/12 PASS`.

## Conclusion

The final 620px PNG sample records normal/managed position P95 values of `524.962/488.881 ms`, a delta of `-36.081 ms`, and confirms `98/98` target selection events. Managed metadata did not increase position-update P95. Because the absolute `300 ms` gate failed, `PointerInputCovered=false`, and Windows denied the pointer interface, this report still does not claim that mouse dragging passed. Continuous real-pointer drag, resize, reposition, and a Re-edit command click remain separate manual acceptance items.

See the [sanitized evidence](./evidence/word-ui-thread-stress-v1.0.8.json) for the release-facing machine-readable summary. User names, absolute attachment paths, temporary run GUIDs, and process identities have been removed while preserving the source report's key configuration, statistics, gate results, and pointer-coverage boundary.
