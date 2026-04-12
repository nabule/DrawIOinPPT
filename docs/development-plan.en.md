# Detailed Development Plan

[English](./development-plan.en.md) | [中文](./development-plan.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Scope

The first formal target release is a single-host PowerPoint Desktop version. WPS is intentionally out of scope for now.

Required product capabilities:

1. Insert Draw.io diagrams into PowerPoint while preserving vector rendering.
2. Re-edit a diagram after selecting the shape.
3. Support editor configuration through either a local desktop path or a URL.
4. Keep the original Draw.io XML stored with the PowerPoint file.

## 2. Current Status

The repository has completed `M0 - Initial Skeleton`, and the first end-to-end versions of `M1`, `M2`, and `M3` have also been closed out:

- Project structure established
- Development plan established
- Data model and serialization implemented
- PowerPoint add-in host entry point implemented
- Ribbon and selection-monitoring skeleton implemented
- Local draw.io desktop workflow closed end-to-end
- First URL-mode editor flow integrated

## 3. Milestones

### M0: Initial Skeleton

Status: completed

Deliverables:

- Solution and project layout
- COM add-in entry point
- Ribbon button definitions
- Selection-monitoring skeleton
- Metadata envelope and settings model
- Development scripts and documentation

### M1: Metadata MVP

Status: completed, first implementation

Goals:

- Detect add-in metadata when a single shape is selected
- Write `DRAWIO_PPT_ID` to the shape
- Write the metadata envelope into `AlternativeText`
- Rehydrate the metadata envelope from an existing shape

Acceptance criteria:

- When a shape with metadata is selected, the add-in recognizes it correctly and shows the right status
- Metadata can still be read after saving and reopening the PPT

Current progress:

- Writing metadata to a selected single shape
- Clearing shape binding
- Rehydrating metadata from the selected shape and refreshing Ribbon state
- Creating new metadata-bearing shapes through the new-diagram flow
- Saving the envelope into `Presentation.CustomXMLParts`
- Restoring shape metadata from document-level storage through `customXmlPartId` or `diagramId`
- Automatically migrating legacy metadata stored only in `AlternativeText` into `CustomXMLParts`
- Backfilling draw.io `content` metadata into generated SVG for better file-level portability

### M2: SVG Insertion and Replacement

Status: completed, first implementation

Goals:

- Insert externally generated SVG into the current slide
- Replace the visible content of an existing add-in-managed shape
- Preserve position, size, and rotation during updates

Acceptance criteria:

- Inserted diagrams remain vector-rendered in PowerPoint
- Replaced diagrams do not jump in layout

Current progress:

- Inserting new SVG shapes into the current slide
- Replacing existing bound shapes while preserving position, size, rotation, and basic z-order
- Manual refresh for the selected shape
- Polling the `.drawio` file in desktop mode and attempting auto-refresh after save
- Launching the editor by double-clicking a bound shape
- Opening the editor on selection when enabled in settings
- Auto-detecting desktop editor paths from common install locations

### M3: Local draw.io Mode

Status: completed, first implementation

Goals:

- Configure the local editor path
- Generate temporary `.drawio` files
- Launch desktop draw.io
- Write back after close or save

Acceptance criteria:

- The local path can be configured and persisted
- Saving can write back both SVG and metadata

Current progress:

- Configuring and auto-detecting the local desktop editor path
- Generating or reusing `.drawio` working files
- Launching desktop `draw.io` / `diagrams.net`
- Auto-refreshing the SVG shape in PowerPoint after save
- Verifying compatible CLI parameters for desktop SVG export

### M4: URL Mode

Status: completed

Goals:

- Configure a usable diagrams.net URL
- Open the editor embedded or in a popup
- Implement XML load/save channels

Acceptance criteria:

- URL mode opens existing diagrams correctly
- Saving updates the diagram in PowerPoint

Current progress:

- `WebView2` host window integrated
- diagrams.net embed iframe integrated
- Host message flow implemented for `init -> load -> save -> export`
- Writing the SVG and XML returned by the URL editor back into the PowerPoint shape
- Local logging and timeout diagnostics for URL initialization and export
- Testing the current URL configuration directly from the settings dialog
- Shipping `scripts/url-editor-smoke.ps1` as a repeatable URL-mode smoke test

### M5: Auto-Update and Stabilization

Status: completed

Goals:

- File monitoring
- Exception handling
- Logging
- Add-in installation guide
- Regression checklist

Acceptance criteria:

- Exceptions do not damage existing PowerPoint content
- Common operation paths leave traceable logs

Current progress:

- Low-frequency cleanup of orphaned `CustomXMLPart` entries with no remaining shape references
- Verification that legacy `AlternativeText` metadata is copied back into document-level storage on read
- Local log file support for diagnosing URL handshake, export, and host write-back issues
- Unified exception handling, user-facing error messages, and logging at the host entry point
- Installation guide and regression checklist completed

## 4. Recommended Development Order

1. Complete M1 first so metadata recognition and persistence are stable.
2. Then complete M2 to make SVG insertion and replacement reliable.
3. Then implement M3, prioritizing the local editor flow.
4. Finally add M4 URL mode.

## 5. Outputs by Phase

### M1 Outputs

- `Shape` metadata read/write service
- Metadata serialization and version management
- Development verification entry points

### M2 Outputs

- SVG insertion service
- SVG replacement service
- Shape-fidelity strategy

### M3 Outputs

- Editor configuration
- Temporary working-directory service
- Local process-launch service
- File write-back orchestration

### M4 Outputs

- Web editor adapter layer
- URL-mode configuration
- Embed API message bridge

### M5 Outputs

- Error reporting
- Debug logs
- Installation scripts and usage instructions

## 6. Risk Breakdown

### High Risk

- Object properties being lost after SVG replacement
- Unstable timing around external-editor save behavior

Mitigation already in place:

- Large XML payloads no longer rely only on `AlternativeText`; they are now stored in document-level `CustomXMLParts`

### Medium Risk

- Different PowerPoint versions handling SVG differently
- Differences in local draw.io launch parameters

### Low Risk

- Settings persistence
- Ribbon status refresh

## 7. Version Suggestions

- `v0.1.0-initial`: repository initialization and host skeleton
- `v0.2.0-metadata`: metadata closed loop
- `v0.3.0-svg`: SVG insertion and replacement
- `v0.4.0-desktop-editor`: local draw.io integration
- `v0.5.0-url-editor`: URL mode
- `v0.6.0-stable`: stabilization and release preparation
