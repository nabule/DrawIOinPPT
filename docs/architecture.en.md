# Architecture

[English](./architecture.en.md) | [中文](./architecture.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Goals

The goal of this project is to give Draw.io diagrams inside PowerPoint the following capabilities:

- Display as SVG while preserving vector quality
- Re-enter Draw.io editing after the shape is selected
- Configure the editor as either a local desktop program or a web URL
- Keep the original Draw.io XML moving with the PPT file as much as possible

## 2. Technical Direction

This project chooses a **native PowerPoint COM Add-in** for these reasons:

- It can reliably listen to PowerPoint shape-selection events
- It can launch local processes directly, which fits desktop draw.io integration
- It handles local paths, temporary files, and later write-back more naturally

WPS is out of scope for now, and Office Web Add-ins are not the primary direction.

## 3. Module Breakdown

### 3.1 `DrawioPpt.Core`

Responsibilities:

- Define the add-in settings model
- Define the Draw.io metadata envelope model
- Provide metadata serialization and compression
- Generate sidecar file paths

The current settings model includes editor mode, desktop editor path, URL editor address, Office-compatible SVG labels, auto-open-on-selection, auto-refresh-after-save, sidecar retention, and whether to show diagram information when creating or editing.

### 3.2 `DrawioPpt.PowerPointAddIn`

Responsibilities:

- COM add-in entry point and registration
- Ribbon command exposure
- Selected-shape monitoring
- PowerPoint object model interaction
- External-editor launch and write-back orchestration

## 4. Data Storage Strategy

### Initial Version Strategy

To reach a usable MVP quickly, the initial version uses two storage layers:

- Shape identity: `Shape.Tags["DRAWIO_PPT_ID"]`
- Shape metadata envelope: `Shape.AlternativeText`

`AlternativeText` stores a custom XML envelope defined by the add-in, containing:

- diagram id
- display name
- editor mode
- editor target
- sidecar path
- updated time
- original draw.io XML

The original draw.io XML is compressed using `gzip + base64` first to reduce size.

### Later Enhancement Strategy

If later validation shows that `AlternativeText` is not sufficient in capacity or compatibility, the second-stage enhancement path is:

- Move the full XML into document-level custom parts or OOXML-attached parts
- Keep only the diagram id and required summary information on the shape itself

## 5. Editing Workflow

### 5.1 Create

1. Click `New Diagram` on the Ribbon
2. Launch the external draw.io editor
3. The user saves the `.drawio` file
4. The add-in exports SVG
5. The add-in inserts the SVG into the current slide
6. The add-in writes `DRAWIO_PPT_ID` and the metadata envelope
7. If `ShowDiagramInfoDialog` is enabled, it shows the diagram name, Diagram ID, editor mode, and `.drawio` working file path

### 5.2 Edit

1. The user selects a shape
2. The add-in listens for the selection change
3. It checks whether the shape is a Draw.io diagram managed by the add-in
4. It reads the metadata envelope
5. It opens the local editor or URL editor
6. The user saves
7. The add-in regenerates SVG and replaces the displayed content
8. If `ShowDiagramInfoDialog` is enabled, it shows the current diagram information when entering the edit flow

## 6. Update Strategy

The first stage updates diagrams by replacing shape content, prioritizing an end-to-end working flow.

The following data is preserved:

- position
- size
- rotation
- z-order

If later validation shows that animations, hyperlinks, or complex formatting are lost too easily during replacement, the enhancement path is:

- Replace the internal picture part inside the PPTX directly to reduce object reconstruction

## 7. Risks and Responses

### Risk 1: `AlternativeText` capacity and compatibility

Response:

- Compress XML first
- Keep the document-level storage upgrade path in the plan

### Risk 2: Detecting saves from the external editor

Response:

- In local mode, prioritize write-back after process exit first
- Add file monitoring and throttling in the second stage

### Risk 3: SVG replacement fidelity

Response:

- Preserve position and size first
- Then address animation and complex-style issues with targeted fixes
