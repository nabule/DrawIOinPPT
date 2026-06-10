# Architecture

[English](./architecture.en.md) | [中文](./architecture.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. Goals

The goal of this project is to give Draw.io diagrams inside Office documents the following capabilities. PowerPoint Desktop is already covered, and this branch adds a Word Desktop host:

- Display as SVG while preserving vector quality
- Re-enter Draw.io editing after a PowerPoint shape or Word picture is selected
- Configure the editor as either a local desktop program or a web URL
- Keep the original Draw.io XML moving with the `pptx` / `docx` file as much as possible

## 2. Technical Direction

This project chooses **native Office COM Add-ins**, currently with PowerPoint and Word hosts, for these reasons:

- It can reliably listen to Office Desktop selection events
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

### 3.3 `DrawioPpt.WordAddIn`

Responsibilities:

- Word COM add-in entry point and registration
- Ribbon command exposure
- Selected-picture monitoring in Word
- Word `InlineShape` / floating `Shape` object model interaction
- Word document-level `CustomXMLParts` storage and cleanup
- External-editor launch and write-back orchestration

The current Word host reuses public editor and SVG services from the PowerPoint assembly. These shared services can later be extracted into a dedicated `OfficeShared` project.

## 4. Data Storage Strategy

### Initial Version Strategy

To reach a usable MVP quickly, the initial version uses two storage layers:

PowerPoint:

- Shape identity: `Shape.Tags["DRAWIO_PPT_ID"]`
- Shape metadata envelope: `Shape.AlternativeText`

Word:

- Picture metadata envelope: `InlineShape/Shape.AlternativeText`
- Diagram display name: `InlineShape/Shape.Title`

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

PowerPoint has now been enhanced to `Presentation.CustomXMLParts + Shape.Tags + Shape.AlternativeText`; Word uses `Document.CustomXMLParts + InlineShape/Shape.AlternativeText`.

If later validation shows that `AlternativeText` is not sufficient in capacity or compatibility, the next enhancement path is:

- Move the full XML into document-level custom parts or OOXML-attached parts
- Keep only the diagram id and required summary information on the shape itself

## 5. Editing Workflow

### 5.1 Create

1. Click `New Diagram` on the Ribbon
2. Launch the external draw.io editor
3. The user saves the `.drawio` file
4. The add-in exports SVG
5. The PowerPoint add-in inserts the SVG into the current slide; the Word add-in inserts it at the current cursor position
6. The add-in writes the metadata envelope and document-level `CustomXMLParts`
7. If `ShowDiagramInfoDialog` is enabled, it shows the diagram name, Diagram ID, editor mode, and `.drawio` working file path

### 5.2 Edit

1. The user selects a shape or picture
2. The add-in listens for the selection change
3. It checks whether the shape is a Draw.io diagram managed by the add-in
4. It reads the metadata envelope
5. It opens the local editor or URL editor
6. The user saves
7. The add-in regenerates SVG and replaces the displayed content
8. If `ShowDiagramInfoDialog` is enabled, it shows the current diagram information when entering the edit flow

## 6. Update Strategy

The first stage updates diagrams by replacing the displayed object, prioritizing an end-to-end working flow.

The following data is preserved:

- position
- size
- rotation
- z-order

In Word, `InlineShape` is part of flowing document text and does not have the same fixed canvas and z-order model as PowerPoint. The current Word path prioritizes preserving picture size, insertion position, and floating-picture wrapping/relative position.

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

### Risk 4: Word flowing layout and floating-picture anchors

Response:

- Prioritize the `InlineShape` picture workflow for the MVP
- Preserve anchor, size, wrapping mode, and relative position when replacing floating `Shape` pictures
- Validate save, reopen, and edit-again behavior with a real Word E2E test
