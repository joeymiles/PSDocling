# PSDocling Documentation Overview

Welcome to the PSDocling help documentation! These guides will help you understand and use all aspects of the PSDocling document processing system.

## Quick Links

| Guide | Description | Best For |
|-------|-------------|----------|
| [Backend Services](01_Backend_Services.md) | API server and document processor setup | System administrators, API users |
| [Frontend Services](02_Frontend_Services.md) | Web interface usage and configuration | End users, web interface users |
| [File Processing](03_File_Processing.md) | Complete processing workflow | All users |

## Requirements

### Minimum Requirements
- Windows PowerShell 5.1 or PowerShell Core 6+
- .NET Framework 4.7.2 (for Windows PowerShell)
- 4GB RAM
- 1GB free disk space

### Recommended Setup
- Windows 10/11 or Windows Server 2016+
- Python 3.8+ with docling package
- 8GB+ RAM for large documents
- SSD for better performance
- Modern browser (Chrome/Edge) for web interface

## Getting Started

### First Time Setup

1. **Install** (builds from source automatically and offers a desktop shortcut)
   ```powershell
   .\app\scripts\Install-DoclingModule.ps1
   ```

2. **Start PSDocling** from the desktop shortcut (or `PSDocling.cmd` in a checkout). It opens in its own window with no console. From PowerShell:
   ```powershell
   Import-Module PSDocling
   Initialize-DoclingSystem
   Start-DoclingSystem -UseWebView
   ```

3. **Process Your First Document**
   - Drag and drop a PDF file into the window
   - Click "Process Documents"
   - Download the converted file

### Choose Your Path

#### I want to use the Web Interface
→ Start with the [Frontend Services Guide](02_Frontend_Services.md)

#### I want to use PowerShell commands
→ Start with the [File Processing Guide](03_File_Processing.md)

#### I want to integrate with the API
→ Start with the [Backend Services Guide](01_Backend_Services.md)

## System Architecture

```
┌─────────────────┐     ┌──────────────────────┐     ┌─────────────────┐
│  App window     │────▶│  API server + UI     │────▶│    Document     │
│  (pywebview)    │     │  localhost:8080 only │     │   Processor     │
└─────────────────┘     └──────────────────────┘     └─────────────────┘
                               │                           │
                               ▼                           ▼
                        ┌──────────────────────────────────────────┐
                        │ %LOCALAPPDATA%\PSDocling\data            │
                        │ queue folder, status.json, uploads, output│
                        └──────────────────────────────────────────┘
```

The API server and processor run as hidden PowerShell processes. Quit (or closing the window) stops all of them.

## Key Features

- **Multiple Input Formats**: PDF, DOCX, XLSX, PPTX, HTML, Markdown, Images
- **Multiple Output Formats**: Markdown, JSON, HTML, DocTags (XML)
- **Enrichment Options**: Code understanding, formula detection, image analysis
- **Web Interface**: Modern drag-and-drop interface with real-time updates
- **REST API**: Full programmatic control
- **Queue-Based**: Reliable asynchronous processing
- **Cross-Platform**: Works on Windows PowerShell and PowerShell Core

## Common Use Cases

### Personal Document Management
```powershell
# Convert all PDFs in a folder to Markdown
Get-ChildItem "C:\MyDocuments" -Filter "*.pdf" | ForEach-Object {
    Add-DocumentToQueue -Path $_.FullName
}
```

### Automated Document Pipeline
```powershell
# Watch folder for new documents
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = "C:\IncomingDocuments"
$watcher.Filter = "*.pdf"
$watcher.EnableRaisingEvents = $true

Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action {
    Add-DocumentToQueue -Path $Event.SourceEventArgs.FullPath
    Write-Host "Auto-processing: $($Event.SourceEventArgs.Name)"
}
```

## Default Ports and Paths

| Component | Default Value | Override |
|-----------|--------------|----------|
| API server and UI | http://localhost:8080 (loopback only) | `-Port` |
| Data folder | %LOCALAPPDATA%\PSDocling | `PSDOCLING_HOME` |
| Queue folder | data\queue | |
| Status file | data\status.json | |
| Uploads | data\uploads | |
| Converted output | data\output | |
| Logs (capped) | logs\ | |
| Write token, process ids | run\ | |

## Quick Troubleshooting

### Service Won't Start
```powershell
# Why the last start failed
Get-Content "$env:LOCALAPPDATA\PSDocling\logs\last-launch-error.txt"

# Check if the port is in use
netstat -ano | findstr :8080

# Use a different port
Start-DoclingSystem -Port 9080 -UseWebView
```

### Documents Not Processing
```powershell
# Check Python availability
Get-PythonStatus

# Run in simulation mode (no Python needed)
Initialize-DoclingSystem -SkipPythonCheck
Start-DoclingSystem -UseWebView
```

### Window Shows Nothing
```powershell
# Restart everything
Stop-DoclingSystem
Start-DoclingSystem -UseWebView

# If the UI files are missing, reinstall
.\app\scripts\Install-DoclingModule.ps1 -Force
```

## Getting Help

### In PowerShell
```powershell
# Get help for any function
Get-Help Add-DocumentToQueue -Full
Get-Help Start-DoclingSystem -Examples
Get-Help Initialize-DoclingSystem -Detailed
```

### Check System Status
```powershell
# Comprehensive status check
Get-DoclingSystemStatus | Format-List *
```

### Debug Information
```powershell
# View logs (api.log, processor.log, processor-errors.log, launcher.log, window.log)
Get-ChildItem "$env:LOCALAPPDATA\PSDocling\logs"
Get-Content "$env:LOCALAPPDATA\PSDocling\logs\processor-errors.log" -Tail 50

# Check queue status
Get-QueueItems | Format-Table -AutoSize

# View processing status
Get-ProcessingStatus | Format-List
```

## Next Steps

1. **Read the appropriate guide** based on your needs
2. **Try the examples** in each guide
3. **Experiment with options** to find what works best

## Version Information

- **Current Version**: 3.3.2
- **Module Type**: PowerShell Script Module
- **License**: See LICENSE file
- **Repository**: https://github.com/joeymiles/PSDocling

---

*For the latest updates and issues, visit the [GitHub repository](https://github.com/joeymiles/PSDocling)*

