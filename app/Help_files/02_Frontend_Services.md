# Frontend Guide

## Overview
The PSDocling window is where you upload documents, watch processing and download converted files. It opens in its own window (pywebview, or an Edge/Chrome app window when pywebview is not installed).

## Table of Contents
- [Starting PSDocling](#starting-psdocling)
- [Default Configuration](#default-configuration)
- [Using the Interface](#using-the-interface)
- [Features and Functionality](#features-and-functionality)
- [Quit, Settings and Uninstall](#quit-settings-and-uninstall)
- [Troubleshooting](#troubleshooting)

## Starting PSDocling

- Double-click the **PSDocling** desktop shortcut (or `PSDocling.cmd` in a checkout). No console window appears.
- On first run PSDocling offers to add the desktop shortcut. Choose **Add shortcut** or **Not now**; it will not ask again, and you can add it later from **Settings**.

From PowerShell:

```powershell
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem -UseWebView     # own window
Start-DoclingSystem -OpenBrowser    # browser tab instead (development)
```

## Default Configuration

- **Address**: `http://localhost:8080` (the API serves the UI; loopback only)
- **Auto-refresh**: every 2 seconds
- **Supported formats**: PDF, DOCX, XLSX, PPTX, MD, HTML, CSV, PNG, JPG, JPEG, TIFF, BMP, WEBP
- **Max upload**: 100 MB per file
- **UI files**: `DoclingFrontend\index.html` and `psdocling.ico` next to the module. `New-FrontendFiles -Destination <folder>` copies them if you want to customise the UI.

## Using the Interface

### Main Interface Components

1. **Header**
   - PSDocling title and version
   - Backend Status: shows "Connected" when the service is running
   - **Settings** and **Quit** buttons

2. **Upload Area**
   - Drag and drop files, or click **Choose Files**
   - Multiple files at once

3. **Queue Overview**: counts for Queued, Processing, Completed and Errors

4. **Processing Results**
   - Per-document output format and enrichment options
   - Start, progress, cancel and error details

5. **Processed Files**
   - Download each result, or **Download All**

Messages appear as notifications in the bottom-right corner; confirmations (such as cancelling a conversion) open an in-app dialog. Press Escape or click outside a dialog to cancel it.

### Step-by-Step

```text
1. Start PSDocling from the desktop shortcut
2. Drag files onto the upload area (or click "Choose Files")
3. Pick the output format and any enrichments for each document
4. Start processing and watch the status move from Queued to Processing to Completed
5. Download results from Processed Files
```

## Features and Functionality

### Processing Options Explained

| Option | Description | Use Case |
|--------|-------------|----------|
| **Markdown** | Plain text with formatting | Documentation, notes |
| **JSON** | Structured data format | Data processing, APIs |
| **HTML** | Web-ready format | Web publishing |
| **DocTags** | XML structured format | Data extraction |
| **Code Understanding** | Extract and analyze code blocks | Technical documents |
| **Formula Detection** | Identify mathematical formulas | Scientific papers |
| **Picture Classification** | Categorize images | Document analysis |
| **Picture Description** | Generate image descriptions | Accessibility |
| **Enable Chunking** | Split large documents | RAG and large files |
| **Embed Images** | Include images in output | Self-contained docs |

### Where Results Are Stored

```text
%LOCALAPPDATA%\PSDocling\data\output\
└── 305d7273-145f-4614-80ef-9933cfec0506\
    ├── Test_File.md
    ├── image_001.png
    └── image_002.png
```

In the window, downloads open a native Save dialog.

## Quit, Settings and Uninstall

- **Quit** stops the service, the processor and the window. If documents are still queued or processing, the dialog says so; finished results are kept. Closing the window does the same as Quit.
- **Settings** shows the data folder, **Add desktop shortcut** (when there is none) and **Uninstall PSDocling**.
- **Uninstall** warns that it permanently deletes the queue, history, uploads, converted documents, logs and settings. Optional checkboxes, all off by default, also remove Python packages PSDocling installed: Docling packages, tokenizers, pywebview. Python itself is never removed. Choose **Keep PSDocling** to back out.

## Troubleshooting

1. **"Backend Status: Disconnected"**
   ```powershell
   Get-DoclingSystemStatus
   Stop-DoclingSystem
   Start-DoclingSystem -UseWebView
   ```

2. **PSDocling does not open from the shortcut**
   ```powershell
   Get-Content "$env:LOCALAPPDATA\PSDocling\logs\last-launch-error.txt"
   Get-Content "$env:LOCALAPPDATA\PSDocling\logs\launcher.log" -Tail 20
   ```

3. **Window is blank**: check `logs\window.log`; install window support with `pip install -r requirements-webview.txt`, or reinstall PSDocling with `-Force`.

4. **Actions fail with 403**: the page was loaded from an older run. Close the window and start PSDocling again (the write token changes at every start).

### Browser Console Debugging

With `-OpenBrowser`, press F12:
- Network tab for API calls
- Console for JavaScript errors

```javascript
fetch('/api/health').then(r => r.json()).then(console.log)
```

## Performance Tips

1. **File sizes**: best under 10 MB; enable chunking for large files
2. **Batches**: upload several small files rather than one very large one

## Security Notes

- Files are processed locally; nothing is sent to external services unless an enrichment you enable does so
- The service listens on localhost only, and other web pages cannot call it
- Changes (upload, cancel, quit, uninstall) need a per-run token that only the app has
- Original files are never modified
