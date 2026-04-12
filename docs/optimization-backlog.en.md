# Optimization Backlog

[English](./optimization-backlog.en.md) | [中文](./optimization-backlog.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. High Priority

### 1. Continue Improving SVG Replacement Fidelity

- The current implementation already preserves position, size, rotation, name, and basic hyperlinks
- `v0.6.1` already enables `simpleLabels` automatically in URL mode
- Still worth improving further:
  - full animation migration
  - more `ActionSettings` scenarios
  - preservation of more complex formatting and layering

### 2. Automating `simpleLabels` for Desktop Mode

- URL mode can already send `{ "simpleLabels": true }` automatically through the add-in
- Desktop mode still depends on the external draw.io Desktop configuration
- Possible next steps:
  - investigate whether there is a safe and reliable way to write desktop configuration
  - show a one-time guidance flow when the desktop editor is first detected
  - validate SVG label rendering quality further in the export pipeline

### 3. Better Release Installation Experience

- The project already ships a release package, `install.cmd`, and PowerShell installation scripts
- Still recommended:
  - a more intuitive one-click install UI
  - a post-install shortcut entry
  - version-upgrade and overwrite-install prompts

### 4. Real Remote Compatibility Matrix for URL Mode

- The project already has `Test URL` and a URL smoke test
- Still recommended:
  - a verification matrix for official `embed.diagrams.net`
  - compatibility notes for private deployments
  - clearer handling around proxy, certificate, and cross-origin issues

## 2. Medium Priority

### 5. Smarter Sidecar Path Relocation

- Sidecar behavior is already designed to follow the PPT in the same directory
- Still recommended:
  - clearer prompts when the sidecar is missing
  - smarter recovery when the sidecar cannot be found

### 6. Batch Operations

- The current workflow is still centered on single-shape operations
- Potential additions:
  - batch refresh
  - batch rebind
  - batch inspection for broken shapes

### 7. Diagnostics Panel

- Troubleshooting currently relies mainly on log files
- Potential additions:
  - `Open log folder` from settings
  - a summary of the latest error
  - per-document checks for shapes and `CustomXMLPart` state

## 3. Low Priority

### 8. Diagram Management View

- Show all managed diagrams in the current presentation
- Filter by `diagramId`, slide number, or update time

### 9. XML Recovery Tool

- When shape metadata is damaged
- Attempt recovery from SVG `content` metadata

### 10. User-Experience Details

- Friendlier prompt text
- Clearer settings explanations
- Better status prompts before and after key actions

## 4. Not in the Current Phase

### 11. WPS Support

- The current version explicitly targets PowerPoint Desktop only
- If WPS support is considered later, it should be planned as a separate host-adapter effort

### 12. Cloud Sync and Collaboration

- Shared storage and multi-user conflict handling are not part of the current product
- This would expand the product boundary significantly and should be planned separately later
