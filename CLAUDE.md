# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Development Commands

### Building the Module
```powershell
# Build module from source files in Source/ directory
.\Build-PSDoclingModule.ps1

# Build and install module
.\Build-PSDoclingModule.ps1 -Install

# Build, install, and test
.\Build-PSDoclingModule.ps1 -Install -Test
```

The build process combines all private and public functions from `Source/` into a single `Build/PSDocling.psm1` file.

### Running the System
```powershell
# Start all services (uses built module if available, falls back to source)
.\Start-All.ps1 -GenerateFrontend -OpenBrowser

# Stop all services
.\Stop-All.ps1

# Custom ports
.\Start-All.ps1 -ApiPort 9080 -WebPort 9081

# Simulation mode (without Python/Docling)
.\Start-All.ps1 -SkipPythonCheck
```

### Module Installation
```powershell
# Install module for current user
.\Install-DoclingModule.ps1

# Install for all users (requires admin)
.\Install-DoclingModule.ps1 -Scope AllUsers

# Uninstall module
.\Uninstall-DoclingModule.ps1
```

## Architecture Overview

PSDocling is a multi-process PowerShell-based document conversion system that wraps Python's Docling library. It uses a queue-based architecture for reliable asynchronous processing.

### Three-Process Architecture

1. **API Server** (`Start-APIServer`)
   - HTTP listener on port 8080 (default)
   - REST endpoints for document upload, status queries, file downloads
   - Runs in separate PowerShell process
   - Uses `System.Net.HttpListener` for HTTP handling
   - Handles CORS for web frontend communication

2. **Document Processor** (`Start-DocumentProcessor`)
   - Background service that monitors queue and processes documents
   - Runs in separate PowerShell process
   - Spawns Python subprocesses for Docling conversion
   - Handles progress tracking, timeouts (6 hours), and cancellation
   - Manages enrichment options (code, formulas, images)
   - Supports hybrid chunking for RAG applications

3. **Web Frontend** (`New-FrontendFiles` + `Start-WebServer.ps1`)
   - Simple HTTP server on port 8081 (default)
   - Single-page application with drag-drop file upload
   - Real-time status updates
   - Generated dynamically by `New-FrontendFiles` function

### Inter-Process Communication

- **Queue File**: `$env:TEMP\docling_queue.json` - Documents waiting for processing
- **Status File**: `$env:TEMP\docling_status.json` - Current state of all documents
- **File Mutex**: Used by `Use-FileMutex` to prevent race conditions on shared JSON files

### Key Data Flow

1. Document uploaded via API → Added to queue → Status set to "Queued"
2. Processor picks up queued item → Status set to "Processing" → Spawns Python
3. Python converts document using Docling → Writes output file → Returns JSON result
4. Processor updates status to "Completed" → Runs optional enhancements (chunking)
5. Frontend polls API for status → Downloads completed file

## Code Organization

### Source Directory Structure
```
Source/
├── Private/           # Internal helper functions
│   ├── Initialize-Module.ps1        # Module initialization code
│   ├── Use-FileMutex.ps1            # Thread-safe file access
│   ├── Get-SecureFileName.ps1       # Security validation
│   └── Test-*.ps1                   # Validation helpers
└── Public/           # Exported module functions
    ├── Configuration/  # Get/Set-DoclingConfiguration
    ├── Processing/     # Document conversion functions
    ├── Queue/          # Queue management (Add/Get/Update)
    ├── Services/       # API/Processor/Frontend services
    └── System/         # Initialization and status
```

### Build Output
- `Build/PSDocling.psm1` - Monolithic module file (all functions combined)
- `Build/PSDocling.psd1` - Module manifest (copied from root)

### Important: Source vs Built Module
- `Start-All.ps1` checks for `Build/PSDocling.psm1` first, falls back to source
- When making changes, rebuild with `.\Build-PSDoclingModule.ps1`
- Module stores state in `$script:DoclingSystem` hashtable

## Key Implementation Details

### Python Integration
The document processor embeds a Python script (in `Start-DocumentProcessor.ps1` lines 69-374) that:
- Imports Docling library
- Configures pipeline options (enrichments, image extraction)
- Converts document to requested format (markdown, HTML, JSON, text, doctags)
- Returns JSON result with success status and image metadata
- Handles image extraction with two modes: embedded (base64) or separate files

### Security Features
- `Get-SecureFileName`: Validates filenames against path traversal attacks
- `Test-SecureFileName`: Validates allowed extensions
- File size limits (100MB default) enforced in API upload endpoint
- Path resolution checks to prevent directory traversal

### Progress Tracking
- File-based estimation for standard conversions (30s to 5min based on size)
- AI enrichment progress models:
  - Picture Description (Granite Vision): 5 min model load + 25 min processing
  - Code/Formula Understanding: 3 min model load + 17 min processing
- Progress never exceeds 95% until completion (prevents user confusion)
- Updates sent every 1% change to avoid excessive file I/O

### Enrichment Options
- **Code Understanding**: `do_code_enrichment` - Analyzes code blocks
- **Formula Detection**: `do_formula_enrichment` - Extracts mathematical formulas
- **Picture Classification**: `do_picture_classification` - Classifies image types
- **Picture Description**: `do_picture_description` - Generates image descriptions using Granite Vision model
- **Hybrid Chunking**: Advanced RAG chunking with semantic and structure-aware splitting

### Output Formats
- **markdown**: `.md` files with optional image references
- **html**: `.html` files with optional image embedding
- **json**: `.json` with full document structure
- **text**: `.txt` plain text extraction
- **doctags**: `.xml` with proper XML wrapping (Docling's native output is malformed)

### Image Handling
- Raster images extracted as PNG files in same directory as output
- Vector graphics detected but not extractable (placeholder inserted)
- Base64 embedding supported for all formats
- Image references use relative paths (just filename)

## Testing Without Python

Use simulation mode for UI and workflow testing:
```powershell
.\Start-All.ps1 -SkipPythonCheck
```
This sets `$script:DoclingSystem.PythonAvailable = $false` and generates mock output files.

## Common Development Tasks

### Adding a New Public Function
1. Create file in `Source/Public/<Category>/New-Function.ps1`
2. Add function name to `FunctionsToExport` in `PSDocling.psd1`
3. Rebuild: `.\Build-PSDoclingModule.ps1`

### Modifying API Endpoints
Edit `Source/Public/Services/Start-APIServer.ps1` - uses regex-based routing in switch statement.

### Changing Queue Behavior
Edit queue functions in `Source/Public/Queue/` - all use `Use-FileMutex` for thread safety.

### Debugging Services
Services run in hidden windows. To debug:
1. Copy script content from `Start-DoclingSystem`
2. Run manually in visible PowerShell window
3. Or check temp files: `$env:TEMP\docling_output.txt`, `$env:TEMP\docling_error.txt`

## File Locations at Runtime

- **Temp Directory**: `$env:TEMP\DoclingProcessor\<guid>\` - Uploaded files during processing
- **Output Directory**: `.\ProcessedDocuments\<document-id>\` - Completed conversions
- **Temp Scripts**: `$env:TEMP\docling_api.ps1`, `$env:TEMP\docling_processor.ps1` - Generated subprocess scripts
- **Frontend Files**: `.\DoclingFrontend\` - Generated by `New-FrontendFiles`

## Module State Management

The module maintains state in `$script:DoclingSystem` hashtable:
```powershell
@{
    Version = '3.1.0'
    Initialized = $true
    PythonAvailable = $true/$false
    APIPort = 8080
    WebPort = 8081
    TempDirectory = "$env:TEMP\DoclingProcessor"
    OutputDirectory = ".\ProcessedDocuments"
    QueuePath = "$env:TEMP\docling_queue.json"
    StatusPath = "$env:TEMP\docling_status.json"
}
```

Access via: `Get-DoclingSystemStatus`

## Dependencies

- **PowerShell**: 5.1+ or Core 6+ (both supported)
- **.NET Framework**: 4.7.2+ (Windows PowerShell only)
- **Python**: 3.8+ with `docling` package (for actual conversion)
- **Docling**: Auto-installed by Python script if missing
- **System.Net.HttpListener**: For API and web servers
- **Compress-Archive**: For ZIP download functionality
