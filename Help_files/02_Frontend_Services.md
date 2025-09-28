# Frontend Web Interface Guide

## Overview
The PSDocling frontend provides a modern web interface for uploading documents, monitoring processing status, and downloading converted files. This guide covers using the web interface with both default and custom configurations.

## Table of Contents
- [Starting the Web Interface](#starting-the-web-interface)
- [Default Configuration](#default-configuration)
- [Custom Domain and Ports](#custom-domain-and-ports)
- [Using the Interface](#using-the-interface)
- [Features and Functionality](#features-and-functionality)
- [Troubleshooting](#troubleshooting)

## Starting the Web Interface

### Quick Start

```powershell
# Start all services with frontend and open browser
.\Start-All.ps1 -GenerateFrontend -OpenBrowser
```

### Manual Start

```powershell
# Import module and initialize
Import-Module PSDocling
Initialize-DoclingSystem -GenerateFrontend
Start-DoclingSystem -OpenBrowser

# Access the interface
Start-Process "http://localhost:8081"
```

## Default Configuration

### Default Settings
- **Web URL**: `http://localhost:8081`
- **API Connection**: `http://localhost:8080`
- **Auto-refresh**: Every 1 second for status updates
- **Supported Formats**: PDF, DOCX, XLSX, PPTX, MD, HTML, CSV, PNG, JPG, JPEG, TIFF, BMP, WEBP

### Example: Starting with Defaults

```powershell
# 1. Generate frontend files and start services
.\Start-All.ps1 -GenerateFrontend -OpenBrowser

# 2. The browser will open automatically
# If not, navigate to: http://localhost:8081

# 3. Verify connection status
# Look for "Backend Status: Connected" in Cisco blue color
```

### Frontend File Generation

```powershell
# The frontend files are generated dynamically
New-FrontendFiles

# Files created in .\DoclingFrontend\:
# - index.html (main interface)
# - styles.css (styling)
# - Start-WebServer.ps1 (web server script)

# View generated files
Get-ChildItem .\DoclingFrontend
```

## Custom Domain and Ports

### Configuration for Custom Ports

```powershell
# Method 1: Using Start-All script
.\Start-All.ps1 -ApiPort 9080 -WebPort 9081 -GenerateFrontend -OpenBrowser

# Method 2: Manual configuration
$config = Get-DoclingConfiguration
$config.Server.APIPort = 9080
$config.Server.WebPort = 9081
Set-DoclingConfiguration -Config $config

# Regenerate frontend with new ports
Initialize-DoclingSystem -GenerateFrontend
Start-DoclingSystem
```

### Setting Up Custom Domain

```powershell
# 1. Update configuration for production domain
@{
    Server = @{
        APIHost = 'api.company.com'
        APIPort = 443
        WebHost = 'app.company.com'
        WebPort = 443
        Protocol = 'https'
        EnableCORS = 'true'
        AllowedOrigins = @('https://app.company.com')
    }
    Frontend = @{
        Title = 'Company Document Processor'
        Theme = 'corporate'
        MaxFileSize = '50MB'
    }
} | Export-PowerShellDataFile -Path ".\PSDocling.config.psd1"

# 2. Regenerate frontend with custom configuration
Initialize-DoclingSystem -GenerateFrontend

# 3. Deploy to web server (IIS example)
Copy-Item .\DoclingFrontend\* C:\inetpub\wwwroot\docling -Force
```

### Example: Corporate Network Setup

```powershell
# Setup for internal corporate network
# Run as Administrator

# 1. Configure for internal network
$internalHost = "docling.internal.corp"
$apiPort = 8080
$webPort = 80

# 2. Update hosts file (optional for testing)
Add-Content -Path C:\Windows\System32\drivers\etc\hosts `
    -Value "127.0.0.1 $internalHost"

# 3. Configure and start
.\Start-All.ps1 -ApiPort $apiPort -WebPort $webPort `
    -GenerateFrontend -EnsureUrlAcl

# 4. Share URL with team
Write-Host "Document processor available at: http://${internalHost}:${webPort}"
```

## Using the Interface

### Main Interface Components

1. **Header Section**
   - Title: "PSDocling" with version (v3.0.0 in smaller text)
   - Subtitle: "PowerShell-based Document Processor"
   - Backend Status: Shows "Connected" in Cisco blue when online

2. **Upload Area**
   - Drag and drop files directly
   - Click "Choose Files" button to browse
   - Supports multiple file selection

3. **Processing Options**
   - Output Format: Markdown, JSON, HTML, DocTags
   - Enrichment Options:
     - Code Understanding
     - Formula Detection
     - Picture Classification
     - Picture Description
   - Advanced Options:
     - Enable Chunking
     - Embed Images

4. **Current Queue**
   - Shows files being processed
   - Real-time progress indicators
   - Status updates (Queued → Processing → Completed)

5. **Processed Files**
   - List of completed documents
   - Download links for each format
   - Re-process option for different formats

### Step-by-Step Upload Process

```text
1. Open browser to http://localhost:8081

2. Upload documents:
   Option A: Drag files onto the upload area
   Option B: Click "Choose Files" and select documents

3. Configure processing options:
   - Select output format (default: Markdown)
   - Enable enrichments if needed
   - Choose chunking for large documents

4. Click "Process Documents"

5. Monitor progress:
   - Watch the progress bar
   - Status changes from "Queued" to "Processing" to "Completed"

6. Download results:
   - Click download link in Processed Files section
   - Files are saved with original name + new extension
```

## Features and Functionality

### File Upload Methods

#### Drag and Drop
```text
1. Select files in Windows Explorer
2. Drag files over the upload area
3. Drop when area highlights
4. Files appear in queue automatically
```

#### Browse and Select
```text
1. Click "Choose Files" button
2. Navigate to document location
3. Select one or multiple files (Ctrl+Click)
4. Click "Open" to add to queue
```

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
| **Enable Chunking** | Split large documents | Large file processing |
| **Embed Images** | Include images in output | Self-contained docs |

### Real-Time Status Updates

The interface automatically updates every second:
- Queue status refreshes
- Progress bars update
- Completed files appear automatically
- Connection status monitored

### Download Management

```text
Processed files are organized by:
- Document ID (GUID folder)
- Original filename preserved
- Multiple output formats available
- Images saved in same folder (not subfolder)

Example structure:
ProcessedDocuments/
├── 305d7273-145f-4614-80ef-9933cfec0506/
│   ├── Test_File.md
│   ├── image_001.png
│   └── image_002.png
```

## Advanced Features

### Batch Processing

```javascript
// The interface supports multiple file selection
// All files process sequentially
// Each maintains individual progress tracking
```

### Custom Styling

```css
/* Cisco blue theme applied to:
- Connected status (#00bceb)
- Progress bars
- Success messages

/* Version number styling:
- 60% of title size
- Light gray color (#999)
- Lighter font weight
*/
```

### Error Handling

The interface provides detailed error information:
- Failed processing shows error message
- Retry option available
- Error details expandable for debugging

## Troubleshooting

### Common Issues

1. **"Backend Status: Disconnected"**
   ```powershell
   # Check if API server is running
   Get-DoclingSystemStatus

   # Restart services
   .\Stop-All.ps1
   .\Start-All.ps1 -GenerateFrontend
   ```

2. **Files Not Uploading**
   ```powershell
   # Check file size limits
   $config = Get-DoclingConfiguration
   Write-Host "Max file size: $($config.Frontend.MaxFileSize)"

   # Check supported formats
   $config.Frontend.AllowedExtensions
   ```

3. **Interface Not Loading**
   ```powershell
   # Regenerate frontend files
   New-FrontendFiles

   # Check if web server is running
   Get-Process | Where-Object {$_.CommandLine -like "*Start-WebServer*"}
   ```

### Browser Console Debugging

Press F12 in browser and check:
- Network tab for API calls
- Console for JavaScript errors
- Check for CORS issues if using custom domains

### Connection Testing

```javascript
// In browser console, test API connection:
fetch('http://localhost:8080/api/health')
  .then(r => r.json())
  .then(console.log)
  .catch(console.error)
```

## Performance Tips

1. **Optimal File Sizes**
   - Best performance: Files under 10MB
   - Large files: Enable chunking
   - Batch processing: Upload multiple small files

2. **Browser Compatibility**
   - Best: Chrome, Edge (Chromium)
   - Good: Firefox
   - Limited: Internet Explorer (not recommended)

3. **Network Considerations**
   - Local processing fastest
   - Network latency affects status updates
   - Large files may timeout over slow connections

## Security Notes

- Files processed locally by default
- No external API calls unless configured
- Temporary files cleaned automatically
- Original files never modified
- CORS protection enabled