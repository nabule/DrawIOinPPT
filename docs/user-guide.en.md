# DrawioPpt User Guide

[English](./user-guide.en.md) | [中文](./user-guide.md) | [English Home](../README.en.md) | [中文首页](../README.md)

## 1. What This Is

DrawioPpt is a native PowerPoint Desktop add-in that brings Draw.io / diagrams.net diagrams into presentations in a way that keeps them editable, writable back to the original shape, and stored with the PowerPoint file as much as possible.

It is designed to solve these core problems:

- Insert SVG vector diagrams into PowerPoint instead of low-quality screenshots
- Keep editing the same diagram later in Draw.io
- Write changes back into the original shape instead of manually replacing it every time
- Preserve the source XML inside the PPT whenever possible so the diagram does not become visually present but no longer editable

Recommended companion docs:

- [Installation and local debugging guide](./installation.en.md)
- [Full E2E test report](./e2e-test-report-v1.0.0.en.md)
- [Release notes for v1.0.0](./release-notes-v1.0.0.en.md)
- [Optimization backlog](./optimization-backlog.en.md)

## 2. What the Current Version Can Do

The current version can already do the following:

- Create a new Draw.io diagram in the current slide
- Insert the diagram as SVG and keep vector rendering quality
- Detect shapes already managed by the add-in
- Re-edit bound shapes
- Support desktop editing with draw.io / diagrams.net Desktop
- Support URL-based editing
- Write the updated SVG and XML back to the original PowerPoint shape after save
- Store source XML in `Presentation.CustomXMLParts`
- Read legacy metadata stored in `AlternativeText` and migrate it automatically during reads
- Backfill draw.io `content` metadata into exported SVG
- Write logs to help diagnose URL-mode and write-back failures
- Group the Ribbon into `Create / Current Shape / Workflow / Info`
- Relocate sidecar `.drawio` files automatically after save-as, path changes, or moving the deck to another machine

Areas still being improved:

- Fidelity for animations, complex actions, and a small number of formatting edge cases after SVG replacement
- Overall URL-mode stability, which still depends on the target deployment and embed-protocol support
- Recovery guidance when a sidecar file is missing after relocation
- Advanced features such as batch refresh, batch rebind, and a diagram management list

## 3. What To Do First After Installation

Recommended first-time setup order:

1. Confirm that PowerPoint has loaded the add-in correctly.
2. Open `Settings` on the add-in Ribbon.
3. Decide whether you want to use `Desktop` mode or `URL` mode first.
4. If you use desktop mode, check the desktop editor path first.
5. If you use URL mode, enter the editor URL and run `Test URL`.
6. Decide whether to enable auto-open when selecting a bound shape.
7. Decide whether to keep sidecar working files next to the PPT.

If the add-in is not installed yet, start with [installation.en.md](./installation.en.md).

## 4. Ribbon Overview

The add-in exposes these main entry points on the PowerPoint Ribbon:

- `New`
  Inserts a new Draw.io diagram into the current slide and enters the editing flow based on the current settings.
- `Edit` or `Re-edit`
  Opens the currently selected Draw.io-managed shape.
- `Refresh`
  Regenerates the current shape from its bound source content.
- `Bind`
  Brings a normal shape under DrawioPpt management.
- `Clear Binding`
  Removes DrawioPpt management from the current shape while keeping its visible appearance on the slide.
- `Auto Open`
  When enabled, selecting a bound shape opens it for editing automatically.
- `Settings`
  Opens the add-in settings dialog.

The right side of the Ribbon also shows status information, such as:

- Whether a shape is currently selected
- Whether the selected shape is a normal shape or a recognized Draw.io shape
- Whether the current editor mode is desktop or URL
- Whether the main edit button is acting as `Edit` or `Re-edit`

## 5. Settings Explained

Every option in the settings window affects the working flow directly. This section explains them one by one.

### 5.1 Editor Mode

UI label: `Editor Mode`

Available values:

- `Desktop`
- `Url`

What it does:

- `Desktop` mode launches the locally installed `draw.io.exe` or `diagrams.net.exe`
- `Url` mode opens an embedded WebView2 editor inside the add-in and communicates through the draw.io embed protocol

Choose desktop mode when:

- You already have draw.io Desktop installed
- You want a more stable local editing experience
- You are on an internal network or do not want to depend on an external URL
- You prefer working with `.drawio` files

Choose URL mode when:

- You do not want to install draw.io Desktop locally
- You already have a stable embed URL
- You want to connect to a privately deployed diagrams.net instance

Recommendation:

- For most individual and enterprise internal-network usage, try `Desktop` first
- Use `Url` when you already know the embed endpoint is available and compatible

### 5.2 Desktop Path

UI label: `Desktop Path`

What it does:

- Points to the local draw.io or diagrams.net executable
- Is only used in `Desktop` mode

Typical values:

- `draw.io.exe`
- `drawio.exe`
- `diagrams.net.exe`

How to fill it:

- Enter the full path manually
- Click `Browse...` and pick the executable
- Click `Auto Detect` and let the add-in search the registry and common install paths

What happens if it is wrong:

- The desktop editor cannot be launched when creating or editing a diagram
- The add-in prompts you to configure the desktop editor path first

Recommendation:

- Try `Auto Detect` first
- If that fails, use `Browse...`

### 5.3 Editor URL

UI label: `Editor URL`

What it does:

- Defines the draw.io / diagrams.net embed page used in URL mode
- Is only used in `Url` mode

Requirements for the URL:

- It must support the draw.io embed protocol
- It must handle messages such as `init`, `load`, `save`, and `export`
- If you use a private deployment, its embed behavior needs to be compatible with the official protocol

Default behavior:

- The add-in provides an `embed.diagrams.net`-style starting URL by default

Important notes:

- Not every diagrams.net page is usable directly
- A page opening successfully does not guarantee save/export callbacks
- The safest path is to use `Test URL` before saving the setting

### 5.4 Use Microsoft Office Compatible SVG Text Labels

UI label: `Use Microsoft Office compatible SVG text labels`

Internal setting:

- `UseOfficeCompatibleSvgLabels`

What it does:

- Primarily affects `Url` mode
- During the URL-editor handshake, the add-in sends `{ "simpleLabels": true }` to draw.io
- The goal is to improve text-edge clarity when SVG is scaled in PowerPoint

When to enable it:

- In almost all cases, keep it enabled
- Especially in URL mode, turning it off is generally not recommended

When it does not apply:

- In `Desktop` mode, the add-in does not rewrite draw.io Desktop's own configuration
- If you want the same effect in desktop mode, configure `simpleLabels` manually in draw.io Desktop

### 5.5 Auto Open Editor When Selecting a Shape

UI label: `Auto open editor when selecting shape`

Internal setting:

- `AutoOpenOnSelection`

What it does:

- When enabled, selecting a bound Draw.io shape attempts to enter edit mode immediately

Good fit when:

- You frequently want a one-click-edit workflow
- You want fewer Ribbon clicks

Less suitable when:

- You often select shapes just for layout work
- You switch quickly across many shapes on the same slide

Recommendation:

- Keep it off while learning the workflow
- Decide later whether it improves your own editing flow

Additional note:

- The add-in includes short suppression logic to avoid instantly reopening an editor right after refresh or save when selection changes again

### 5.6 Auto Refresh Shape After Save

UI label: `Auto refresh shape after save`

Internal setting:

- `AutoUpdateOnSave`

What it does:

- Primarily affects `Desktop` mode
- The add-in monitors the associated `.drawio` working file
- When you save in the desktop editor, the add-in attempts to regenerate and replace the PowerPoint shape automatically

Good fit when:

- You want PowerPoint to update as soon as the `.drawio` file is saved

What happens if you turn it off:

- You can still use the `Refresh` button manually
- The add-in just stops performing automatic write-back after save

Recommendation:

- Usually keep it enabled
- Disable it only if you suspect auto-refresh is getting in the way

### 5.7 Keep Sidecar Working Files Next to the PPT

UI label: `Keep sidecar file next to PPT`

Internal setting:

- `KeepSidecarFile`

What it does:

- Decides whether the add-in stores the `.drawio` working file near the PPT file

When enabled:

- The add-in tries to place the `.drawio` file in a subfolder next to the PPT
- This is better when you want to move the PPT and its working files together

When disabled:

- The add-in is more likely to use the system temp directory
- This is more convenient for temporary decks or one-off edits

Enable it when:

- You want to maintain the PPT over time
- You want to keep the `.drawio` working files
- You want the related working files visible near the PPT

Disable it when:

- You do not want extra sidecar files in the project folder
- The deck is temporary

### 5.8 Sidecar Folder

UI label: `Sidecar Folder`

Internal setting:

- `SidecarFolderName`

What it does:

- When sidecar retention is enabled, this decides which subfolder stores the `.drawio` working files

Default value:

- `drawio`

Example:

If your PPT is:

```text
D:\Project\demo.pptx
```

And the sidecar folder is:

```text
drawio
```

Then the sidecar file will typically look like:

```text
D:\Project\drawio\demo.shape-name.drawio
```

Change it when:

- You want the working files in another subfolder
- Your project already has a naming convention for attachments or generated assets

Keep the default when:

- You just want the standard behavior

### 5.9 What the Browse, Auto Detect, and Test URL Buttons Mean

- `Browse...`
  Lets you select the draw.io Desktop executable manually.
- `Auto Detect`
  Lets the add-in search the registry and common install paths for the desktop editor.
- `Test URL`
  Opens the URL editor and performs a full handshake check without needing to save the setting first.
  This test checks for save and export callbacks and shows the log path for diagnostics.

## 6. Recommended Configurations

If you are not sure how to configure the add-in, start with one of these presets.

### 6.1 Safest Local Configuration

Best for most personal users and enterprise internal-network environments:

- `Editor Mode`: `Desktop`
- `Desktop Path`: use `Auto Detect`
- `Use Microsoft Office compatible SVG text labels`: enabled
- `Auto open editor when selecting shape`: disabled at first
- `Auto refresh shape after save`: enabled
- `Keep sidecar file next to PPT`: enabled
- `Sidecar Folder`: keep the default `drawio`

### 6.2 Lowest-Friction URL Configuration

Best for teams that already have a stable embed endpoint:

- `Editor Mode`: `Url`
- `Editor URL`: set to a working embed URL
- `Use Microsoft Office compatible SVG text labels`: enabled
- `Auto open editor when selecting shape`: personal preference
- `Auto refresh shape after save`: can stay enabled
- `Keep sidecar file next to PPT`: depends on project needs

Always click `Test URL` once before relying on it in daily use.

### 6.3 Temporary-Deck Configuration

Good for quick presentations when you do not want many sidecar files:

- `Editor Mode`: whichever fits your case
- `Auto open editor when selecting shape`: disabled
- `Auto refresh shape after save`: enabled
- `Keep sidecar file next to PPT`: disabled

## 7. Standard Workflows

This section walks through the most common actions.

### 7.1 Create a New Draw.io Diagram

1. Open a PowerPoint presentation.
2. Switch to the add-in Ribbon.
3. Click `New`.
4. The add-in inserts an SVG shape into the current slide first.
5. Then, depending on the current editor mode:
   - `Desktop`: it tries to launch local draw.io
   - `Url`: it opens the embedded URL editor
6. After saving in the editor, the diagram is written back to the current PPT.

Additional notes:

- If desktop mode is selected but the desktop path is not configured correctly, the shape is inserted as a preview first
- If the PPT has not been saved to disk yet and sidecar retention is enabled, the add-in stores the `.drawio` working file in the system temp directory first to avoid invalid relative paths
- Once the PPT is later saved to a real path, the next edit relocates the sidecar automatically to the configured folder next to the PPT

### 7.2 Re-Edit an Existing Diagram

Method 1:

1. Select a shape already managed by the add-in.
2. Click `Edit` or `Re-edit`.

Method 2:

1. Double-click the bound shape directly.

Method 3:

1. Enable `Auto open editor when selecting shape`.
2. After that, selecting a bound shape opens it automatically.

### 7.3 Refresh a Diagram

Useful when:

- You already saved the `.drawio` file in desktop draw.io
- Automatic write-back did not complete in the current round
- You want to trigger a refresh manually

Steps:

1. Select the shape.
2. Click `Refresh`.

### 7.4 Bind a Normal Shape

Useful when:

- The current selected shape was not originally created by DrawioPpt
- You still want it to enter the add-in's management flow

Steps:

1. Select a normal shape.
2. Click `Bind`.
3. The add-in writes Draw.io metadata and internal identifiers onto it.

Important note:

- Binding does not automatically turn an arbitrary shape into a real draw.io-exported SVG
- More precisely, it brings the shape under the add-in's metadata management model

### 7.5 Clear Binding

Steps:

1. Select a bound shape.
2. Click `Clear Binding`.

Result:

- The shape remains on the slide
- The add-in removes its binding metadata
- It is no longer recognized as a Draw.io-managed shape

Useful when:

- You want to keep the visible shape but stop managing it through the add-in

## 8. Difference Between Desktop Mode and URL Mode

### 8.1 Desktop Mode

Advantages:

- Usually a more stable local editing experience
- Better for enterprise internal networks or offline environments
- Better for long-term maintenance of `.drawio` working files

How it works:

- The add-in prepares a `.drawio` file
- It launches local draw.io Desktop
- After you save, the add-in watches for file changes and tries to refresh the PowerPoint shape

What you need to know:

- You must configure a valid desktop editor path
- If sidecar retention is enabled, the `.drawio` file is stored alongside the PPT whenever possible
- If the PPT is not saved yet, the sidecar goes to the system temp directory first
- If the PPT is saved as a new file, moved, or opened on another machine, the next edit rebuilds the sidecar against the current PPT path
- If you want better SVG text scaling in desktop mode, set `simpleLabels` manually in draw.io Desktop

### 8.2 URL Mode

Advantages:

- No dependency on a locally installed desktop draw.io
- Easier to connect to private deployments
- `Test URL` can validate the flow before actual editing

How it works:

- The add-in opens WebView2 internally
- It communicates with the embedded page through the draw.io embed protocol
- On save, it receives XML and SVG and writes them back to the shape

What you need to know:

- `Editor URL` must support the embed protocol
- Network, proxy, certificate, and private-deployment differences can affect the handshake
- The add-in injects `simpleLabels` automatically in URL mode

## 9. How Data Is Stored

The add-in currently uses a layered persistence model.

### 9.1 Primary Storage Inside the PPT

The main source data is written to:

- `Presentation.CustomXMLParts`

This is the most important layer because it keeps the source XML traveling with the PowerPoint file itself whenever possible.

### 9.2 References Stored on the Shape

The shape itself keeps:

- `Shape.Tags`

This is primarily used to store:

- `DRAWIO_PPT_ID`
- `DRAWIO_PPT_PART_ID`

These values help the add-in map the current shape back to the document-level XML part.

### 9.3 Compatibility Fallback

The add-in still supports:

- `Shape.AlternativeText`

If it encounters an older shape that only stores source data there, it tries to migrate that content into `CustomXMLParts` automatically during reads.

### 9.4 Sidecar `.drawio` Working File

If sidecar retention is enabled, the add-in also keeps a `.drawio` file on disk for:

- Direct opening in the desktop editor
- Manual inspection or backup
- File-change-based auto-refresh in desktop mode

### 9.5 Extra Metadata in the SVG

The add-in also backfills draw.io `content` metadata into generated SVG so that even a single SVG file carries some recoverable diagram information.

## 10. Logging and Troubleshooting

Default log path:

```text
C:\Users\<your-username>\AppData\Roaming\Greensoft\DrawioPpt\Logs\drawioppt.log
```

Check the log first when:

- The URL editor does not open
- The URL opens, but save is never called back
- Saving in URL mode does not export SVG
- Desktop mode opens the editor, but the shape does not refresh after save
- Binding or clearing binding reports an exception

A practical troubleshooting order is:

1. Check whether you are in `Desktop` or `Url` mode
2. Verify the key settings for that mode
3. Inspect the log
4. Retry manually after confirming the above

## 11. Frequently Asked Questions

### 11.1 I clicked New, but the editor did not open

Common reasons:

- You are in desktop mode and `Desktop Path` is not valid
- You are in URL mode and `Editor URL` is not valid
- The URL opens, but does not support the embed protocol

Check first:

- In desktop mode, verify `Desktop Path`
- In URL mode, run `Test URL`

### 11.2 I saved the `.drawio` file in desktop mode, but PowerPoint did not update

Check first:

- Whether `Auto refresh shape after save` is enabled
- Whether the `.drawio` file was actually saved successfully
- Whether the log contains a refresh exception

If it still does not work:

1. Return to PowerPoint
2. Select the shape
3. Click `Refresh` manually once

### 11.3 The URL opens, but Test URL still fails

Because "the page opens" and "the page supports the draw.io embed protocol" are not the same thing.

`Test URL` mainly checks:

- Whether `init` is received
- Whether `load` works
- Whether `save` arrives
- Whether `export` arrives

If those callbacks are incomplete, the page may still open, but it is not reliable enough for the add-in workflow.

### 11.4 Why is the `.drawio` file sometimes stored in a temp directory

Usually because:

- The current PPT has not been saved to disk yet
- Or sidecar retention is disabled

This is normal. It avoids generating incorrect relative working-file paths for a presentation that does not yet have a real save location.

Additional note:

- Once the PPT is saved to a real path later, the next edit moves the sidecar back near the current PPT location

### 11.5 Why does text scaling in PowerPoint sometimes look soft

If you use URL mode:

- Keep `Use Microsoft Office compatible SVG text labels` enabled

If you use desktop mode:

- The add-in does not rewrite draw.io Desktop settings automatically
- You need to enable `simpleLabels` manually in draw.io Desktop

### 11.6 Will Clear Binding delete the shape

No.

`Clear Binding` only removes the add-in management relationship. It does not delete the shape from the slide.

### 11.7 If I bind a normal shape, does it become a real Draw.io diagram

Not exactly.

Binding is better understood as bringing the current shape under the add-in's metadata management model. Whether it can then be edited and written back like a Draw.io diagram still depends on whether the source data is complete and whether the current editing flow succeeds.

## 12. Good and Less-Ideal Usage Scenarios

Good fits for the current version:

- PowerPoint documents maintained by a single owner
- Scenarios where Draw.io source data should travel with the PPT
- PowerPoint Desktop environments
- Teams comfortable with script-based installation and light configuration

Use more carefully in scenarios such as:

- Presentations that rely heavily on animations and complex actions
- Many people modifying the same PPT frequently
- Highly managed enterprise environments that require standardized installers and controlled release pipelines

## 13. Easiest Way to Get Started

If you just want the most stable first experience, use this setup:

1. Open `Settings`
2. Set `Editor Mode` to `Desktop`
3. Click `Auto Detect`
4. Keep `Auto refresh shape after save` enabled
5. Keep `Keep sidecar file next to PPT` enabled
6. Leave `Auto open editor when selecting shape` disabled at first
7. Create one new diagram, save once, and confirm that write-back works

Once that flow is stable, then try URL mode and auto-open.
