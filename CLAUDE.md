# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a PowerShell-based document processing system that converts various document formats (PDF, DOCX, PPTX, XLSX, HTML, MD) to Markdown using the Python Docling library. The system consists of three main components: a REST API server, a document processor, and a web frontend.

## Architecture

### Core Components

1. **DoclingSystem.psm1** - Main PowerShell module containing all system functionality
2. **Start-All.ps1** - System startup script with configuration options
3. **Stop-All.ps1** - System shutdown script
4. **DoclingFrontend/** - Web UI directory containing HTML frontend and web server built dynamicly from DoclingSystem.psm1

### System Design

The system uses a queue-based architecture with file-based persistence:

- **Queue Management**: JSON-based queue stored in temp files for cross-process communication
- **Status Tracking**: Centralized status management using JSON files
- **Multi-Process**: API server, document processor, and web server run as separate PowerShell processes
- **Python Integration**: Calls Python Docling library via subprocess for actual document conversion

### Key Functions (DoclingSystem.psm1)

- `Initialize-DoclingSystem` - Sets up directories, queue, and checks Python/Docling availability
- `Start-DoclingSystem` - Launches all three services (API, processor, web server)
- `Add-DocumentToQueue` - Queues documents for processing
- `Start-DocumentProcessor` - Main processing loop that converts documents
- `Start-APIServer` - REST API server handling web requests
- `New-FrontendFiles` - Generates the HTML frontend and web server script

## Common Commands

### System Startup
```powershell
# Basic startup
.\Start-All.ps1

# With frontend generation and browser opening
.\Start-All.ps1 -GenerateFrontend -OpenBrowser

# Custom ports
.\Start-All.ps1 -ApiPort 9080 -WebPort 9081

# Skip Python check (simulation mode)
.\Start-All.ps1 -SkipPythonCheck
```

### System Management
```powershell
# Stop all services
.\Stop-All.ps1

# Check system status
Import-Module .\DoclingSystem.psm1
Get-DoclingSystemStatus

# Add documents to queue
Add-DocumentToQueue -Path "path\to\document.pdf"
```

### Development/Testing
```powershell
# Import module for interactive testing
Import-Module .\DoclingSystem.psm1 -Force

# Initialize without starting services
Initialize-DoclingSystem -GenerateFrontend

# Check Python/Docling availability
Get-PythonStatus

# Manual queue operations
$queue = Get-QueueItems
Add-QueueItem @{Id="test"; FileName="test.pdf"; Status="Queued"}
$next = Get-NextQueueItem
```

## Configuration

### Default Ports
- API Server: 8080
- Web Frontend: 8081

### File Locations
- Queue file: `$env:TEMP\docling_queue.json`
- Status file: `$env:TEMP\docling_status.json`
- Temp directory: `$env:TEMP\DoclingProcessor`
- Output directory: `.\ProcessedDocuments`

### Supported Formats
- **Documents**: PDF, DOCX, XLSX, PPTX, MD, HTML, XHTML, CSV
- **Images**: PNG, JPEG, TIFF, BMP, WEBP

## Development Notes
- index.html is generated from DoclingSystem.psm1 file. When making edits, do not apply edits to index.html as those edits will be overwritten when the index.html file is rebuilt from DoclingSystem.psm1.

### Python Integration
The system requires Python with the `docling` package. It automatically:
- Detects Python availability on startup
- Installs Docling if missing
- Falls back to simulation mode if Python unavailable

### Error Handling
- Comprehensive error capture with stack traces
- Timeout handling (10 minutes max per document)
- Detailed error reporting via API endpoints

### Cross-Process Communication
Uses JSON files for state sharing between PowerShell processes:
- Queue operations are atomic to prevent corruption
- Status updates are immediately persisted
- All processes can read current system state

### Frontend Architecture
Single-page application with:
- File upload via drag-drop or file selection
- Real-time status updates via polling
- Download links for processed documents
- Detailed error reporting with modal dialogs

## Testing

The system includes simulation mode for testing without Python:
```powershell
.\Start-All.ps1 -SkipPythonCheck
```

This creates mock processed documents for UI and workflow testing.