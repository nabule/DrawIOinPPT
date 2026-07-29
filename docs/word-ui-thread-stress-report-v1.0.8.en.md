# v1.0.8 Word UI-thread Stress Test Report

[English](./word-ui-thread-stress-report-v1.0.8.en.md) | [中文](./word-ui-thread-stress-report-v1.0.8.md) | [Sanitized machine-readable evidence](./evidence/word-ui-thread-stress-v1.0.8.json)

## Purpose and Boundary

This test runs in real Word against the standard local installation. It renders the same user-supplied complex Draw.io source as an aspect-preserving `1240 px`-wide PNG, then creates a normal picture and a managed picture with matching visuals, dimensions, anchor semantics, and wrapping. Both are floating `Shape` objects with `Front` wrapping. The script applies the same selection, position, and size property updates on Word's visible STA UI thread to compare observable managed-metadata overhead.

The standard local Word DLL measured in this run has SHA256 `A6AB59AA0CDA6AA19B0D3A09150494BF4DB6419774F376A28986314B8C3492E8`, matching source Release output and the final package.

The test updates `Shape` properties through Word COM; it does not inject a real mouse pointer. `PointerInputCovered=false`. Earlier attempts to use desktop interfaces reported `GetCursorPos=AccessDenied` and `SetIsBorderRequired=Unsupported`. This report is therefore not mouse-drag acceptance and never represents the automated result as “real mouse passed.” The user chose to skip separate manual pointer acceptance, whose status is “skipped/not accepted.”

The Word add-in does not subscribe to `WindowSelectionChange`. An independent counter verifies the target-picture events raised by Word. During normal use, the user selects a picture and then clicks Re-edit or another command; only that command reads the live selection and starts editing. Selection, dragging, or resizing alone never opens the editor.

## Reproducible Command

Close all Word windows first. `<USER_DRAWIO_SOURCE>` denotes a caller-supplied local file; the report and evidence do not retain its absolute attachment path:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 `
  -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" `
  -DrawioSourcePath "<USER_DRAWIO_SOURCE>" `
  -PictureWrapMode Front `
  -PictureRenderFormat Png `
  -PreviewPixelWidth 1240 `
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
  -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-user-source-v1.0.8-1240px.json"
```

The script runs two crossed-order rounds—normal then managed, followed by managed then normal—in a document containing 150 body paragraphs, 19,456 body characters, eight pages, two tables, and three auxiliary pictures. The source contains 66 `mxCell` elements. The PNG is `1240×1753 px`. The script exports a separate SVG probe to obtain an independent source ratio before comparing the PNG and verifies draw.io uses `--border 0`. Transparent-edge scanning requires 100% of each row or column to have alpha no greater than 8, retains a 2px rasterization-fringe allowance, and stops immediately for a 1px edge-touching line. Opaque 24bpp PNG/JPEG files explicitly report `BlankPaddingDetectionSupported=false` and a null `BorderPaddingPassed` instead of claiming a pixel-level pass. The actual PNG is opaque, so current evidence against forced 1:1 sizing and export borders is the independent ratio error of `0.000236` plus `--border 0`. The normal and managed pictures have the same size, anchor semantics, and `Front` wrapping.

## Results

Execution date: 2026-07-29.

| Metric | Normal picture | Managed picture |
| --- | ---: | ---: |
| Combined duration over two rounds | 21.689 s | 20.608 s |
| Selection operations | 48 | 48 |
| Selection P95 | 68.866 ms | 74.243 ms |
| Position updates | 24 | 24 |
| Mean position-update latency | 202.477 ms | 196.560 ms |
| Position-update P95 | 321.844 ms | 325.764 ms |
| Final position-update failures | 0 | 0 |
| Size updates | 24 | 24 |
| Size-update P95 | 89.633 ms | 122.968 ms |
| Final size-update failures | 0 | 0 |
| Word unresponsive samples | 0 | 0 |

- The combined managed/normal position-P95 ratio is `1.012`, with a delta of `3.920 ms`.
- The selection-P95 ratio is `1.078`, with a `5.377 ms` delta. The resize-P95 ratio is `1.372`, with a `33.335 ms` delta. Per-round variation causes the relative gates to report `ComparisonPassed=false` and `NoAdditionalMetadataPathDifferenceObserved=false`.
- The independent counter confirmed all `96` expected target-picture selection events, with zero missing: `96/96` and `SelectionChangeEventsPassed=true`.
- Save and reopen preserved aspect ratio, position, size, and all 26,106 source-XML characters in `Document.CustomXMLParts`; `WordAddInConnect=true`, and cleanup left no residual `WINWORD`.
- Managed-picture round position P95 values were `303.892/337.339 ms`, while the normal control reported `328.256/318.748 ms`; every pair contains values above `300 ms`, so `AbsoluteLatencyGatePassed=false`. An earlier confirmation run passed the relative gates and kept both managed rounds below 300ms, showing timing variation but not grounds to relax the gate.
- The machine result is `ComparisonPassed=false`, `AbsoluteLatencyGatePassed=false`, and `OverallPassed=false`. Both pictures have zero final selection, position, and resize failures and zero Word-unresponsive samples. These stress gates are distinct from the functional full Office E2E result of `12/12 PASS`.

## Conclusion

The final 1240px PNG sample independently validates source ratio and `--border 0`; transparent-padding pixel inspection is explicitly unsupported for the actual opaque PNG. The latest run records normal/managed selection P95 values of `68.866/74.243 ms`, position P95 values of `321.844/325.764 ms`, resize P95 values of `89.633/122.968 ms`, and confirms `96/96` target selection events. Both relative and absolute gates failed, so `OverallPassed=false`. `PointerInputCovered=false`, and the user chose to skip manual mouse acceptance; this report makes no claim that mouse dragging passed.

See the [sanitized evidence](./evidence/word-ui-thread-stress-v1.0.8.json) for the release-facing machine-readable summary. User names, absolute attachment paths, temporary run GUIDs, and process identities have been removed while preserving the source report's key configuration, statistics, gate results, and pointer-coverage boundary.
