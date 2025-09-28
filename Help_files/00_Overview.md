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

1. **Build and Install the Module**
   ```powershell
   .\Build-PSDoclingModule.ps1
   .\Install-DoclingModule.ps1 -Force
   ```

2. **Start Everything**
   ```powershell
   .\Start-All.ps1 -GenerateFrontend -OpenBrowser
   ```

3. **Process Your First Document**
   - Open browser to http://localhost:8081
   - Drag and drop a PDF file
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
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  Web Interface  │────▶│   API Server    │────▶│    Document     │
│  (Port 8081)    │     │  (Port 8080)    │     │   Processor     │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                               │                         │
                               ▼                         ▼
                        ┌─────────────┐          ┌─────────────┐
                        │ Queue File  │          │ Status File │
                        │   (JSON)    │          │   (JSON)    │
                        └─────────────┘          └─────────────┘
```

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

### Team Document Processing Server
```powershell
# Set up shared server for team
.\Start-All.ps1 -ApiPort 80 -WebPort 80 -EnsureUrlAcl
Write-Host "Share this URL with your team: http://$(hostname)"
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

| Component | Default Value | Environment Variable |
|-----------|--------------|---------------------|
| API Server | http://localhost:8080 | - |
| Web Interface | http://localhost:8081 | - |
| Queue File | %TEMP%\docling_queue.json | $env:TEMP |
| Status File | %TEMP%\docling_status.json | $env:TEMP |
| Temp Processing | %TEMP%\DoclingProcessor | $env:TEMP |
| Output Directory | .\ProcessedDocuments | Current directory |

## Quick Troubleshooting

### Service Won't Start
```powershell
# Check if ports are in use
netstat -ano | findstr :8080
netstat -ano | findstr :8081

# Use different ports
.\Start-All.ps1 -ApiPort 9080 -WebPort 9081
```

### Documents Not Processing
```powershell
# Check Python availability
Get-PythonStatus

# Run in simulation mode (no Python needed)
.\Start-All.ps1 -SkipPythonCheck
```

### Can't Access Web Interface
```powershell
# Regenerate frontend files
New-FrontendFiles

# Restart all services
.\Stop-All.ps1
.\Start-All.ps1 -GenerateFrontend -OpenBrowser
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
# View error logs
Get-Content "$env:TEMP\docling_error.txt" -Tail 50

# Check queue status
Get-QueueItems | Format-Table -AutoSize

# View processing status
Get-ProcessingStatus | Format-List
```

## Next Steps

1. **Read the appropriate guide** based on your needs
2. **Try the examples** in each guide
3. **Experiment with options** to find what works best
4. **Check the HowTo.ps1** file for more examples

## Version Information

- **Current Version**: 3.0.0
- **Module Type**: PowerShell Script Module
- **License**: See LICENSE file
- **Repository**: https://github.com/joeymiles/PSDocling

---

*For the latest updates and issues, visit the [GitHub repository](https://github.com/joeymiles/PSDocling)*