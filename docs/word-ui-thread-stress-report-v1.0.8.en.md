# v1.0.8 Word UI-thread Stress Acceptance Report

[English](./word-ui-thread-stress-report-v1.0.8.en.md) | [中文](./word-ui-thread-stress-report-v1.0.8.md) | [Path-normalized evidence snapshot](./evidence/word-ui-thread-stress-v1.0.8.json)

## Purpose and Boundary

This test runs in real Word against the standard local installation. It applies the same UI-thread operation sequence at the same viewport position to a normal SVG picture and a managed Draw.io picture with identical visuals and dimensions. Each sample selects a neutral document range, allows 75 ms for message processing, then times the target-picture selection separately from the position or size update. The measured instance must have the add-in connected. The test's independent counter classifies events by Shape name and resets that Shape after warmup, confirming target-picture `WindowSelectionChange` events without allowing neutral-range events to satisfy the gate. Separate runtime reflection on the release DLL reports `NoPassiveSelectionMetadataPath=True`; the add-in itself does not subscribe to that event.

The test updates `Shape` properties on the visible Word STA UI thread through Word COM. It does not inject real mouse-pointer input, so it is not manual pointer-drag acceptance and cannot prove pixel-by-pixel tracking of a real pointer.

## Reproducible Command

Close all Word windows first:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\word-ui-thread-stress-acceptance.ps1 `
  -InstallRoot "$env:LOCALAPPDATA\Greensoft\DrawioPpt" `
  -DurationSeconds 12 `
  -MaximumP95Ratio 1.25 `
  -MaximumPerRoundP95Ratio 1.50 `
  -MaximumSelectionP95DeltaMs 100 `
  -MaximumAbsoluteSelectionP95Ms 2000 `
  -MaximumAbsoluteP95Ms 2000 `
  -MaximumAbsoluteResizeP95Ms 750 `
  -MinimumOperationsPerRound 10 `
  -WarmupSeconds 4 `
  -ResizeOperationsPerRound 12 `
  -SelectionSettleMilliseconds 75 `
  -OutputPath ".\artifacts\test-reports\word-ui-thread-stress-v1.0.8.json"
```

The script generates 120 SVG elements and 1,200 Draw.io cells. It runs two rounds with crossed order: normal then managed in round 1, and managed then normal in round 2. Each picture receives a separate four-second warmup per round, at least 12 seconds of selection events and position updates, and 12 single-axis size updates. The pictures are shown separately at the same page position during measurement, reducing viewport, first-layout, overlap-rendering, and fixed-order bias.

## Results

Execution date: 2026-07-28.

| Metric | Normal picture | Managed picture |
| --- | ---: | ---: |
| Combined duration over two rounds | 24.184 s | 24.516 s |
| Timed selection events | 66 | 69 |
| Mean selection-event latency | 3.594 ms | 3.368 ms |
| Selection-event P95 | 7.501 ms | 5.833 ms |
| Position updates | 42 | 45 |
| Mean position-update latency | 311.822 ms | 287.630 ms |
| Position-update P95 | 388.910 ms | 368.568 ms |
| Final position-update failures | 0 | 0 |
| Single-axis size updates | 24 | 24 |
| Single-axis size-update P95 | 112.280 ms | 103.247 ms |
| Final size-update failures | 0 | 0 |
| Word unresponsive samples | 0 | 0 |

- Round 1 normal/managed position P95 values are `387.316/368.568 ms` with ratio `0.952`; round 2 values are `388.910/314.240 ms` with ratio `0.808`.
- Combined managed/normal P95 ratios are `0.778` for selection events, `0.948` for position updates, and `0.920` for size updates, all below the `1.25` threshold. The highest per-round selection, position, and size ratios are `1.015`, `0.952`, and `1.151`, all below `1.50`. The combined selection delta is `-1.668 ms`.
- The highest P95 values across the four groups are `8.766 ms` for selection, `388.910 ms` for position, and `129.184 ms` for size, below the absolute `2000 / 2000 / 750 ms` gates. Each position group contains at least `21` samples versus the required minimum of `10`.
- The measured instance reports `MeasurementWordAddInConnect=True`. The test counter observed `414` total events and classified `135` target-Shape events for `135` target selection operations, with zero missing; therefore `SelectionChangeEventsPassed=True`. Installed-build runtime reflection also reports `NoPassiveSelectionMetadataPath=True`.
- Draw.io XML: `254,606` characters; lightweight managed-picture `AlternativeText`: `382` characters.
- Save and reopen preserved position and size, and `Document.CustomXMLParts` restored all `254,606` characters of source XML.
- `SelectionComparisonPassed=True`, `ComparisonPassed=True`, `AbsoluteLatencyGatePassed=True`, `MinimumSamplesPassed=True`, `WordAddInConnect=True`, `CleanupResidualWinWord=False`, and `OverallPassed=True`.

This sample supports only the statement that no additional picture-metadata-path difference above the configured threshold was observed and that the absolute latency remained within this test's acceptance gates. It does not establish the cause of the absolute latency and does not replace a manual continuous 10-second pointer drag, resize, reposition, and Re-edit button click on each normal and managed picture.

See the [path-normalized evidence snapshot](./evidence/word-ui-thread-stress-v1.0.8.json) for the complete machine-readable result. The snapshot changes only the user-directory prefixes in the real output to `%LOCALAPPDATA%` and `%TEMP%`; all other fields match that script run.
