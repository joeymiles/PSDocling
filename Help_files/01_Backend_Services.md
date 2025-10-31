# Backend Services Guide

## Overview
The PSDocling backend services provide REST API endpoints and document processing capabilities. This guide covers starting and managing the backend services with both default and custom configurations.

## Table of Contents
- [Starting Backend Services](#starting-backend-services)
- [Default Configuration](#default-configuration)
- [Custom Domain and Ports](#custom-domain-and-ports)
- [API Endpoints](#api-endpoints)
- [Health Monitoring](#health-monitoring)
- [Stopping Services](#stopping-services)

## Starting Backend Services

### Prerequisites
- PowerShell 5.1 or higher
- Python 3.8+ with docling package (optional for simulation mode)
- Administrator rights (for custom domains/ports)

### Quick Start with Defaults

```powershell
# Import the module and start services
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem
```

Or use the convenience script:
```powershell
.\Start-All.ps1
```

## Default Configuration

### Default Ports and Addresses
- **API Server**: `http://localhost:8080`
- **Document Processor**: Background service (no direct access)
- **Status Files**: Located in `$env:TEMP`
  - Queue Folder: `$env:TEMP\DoclingQueue` (folder-based queue system)
  - Status: `$env:TEMP\docling_status.json`

### Starting Individual Services

```powershell
# Start only the API server
Start-APIServer -Port 8080

# Start only the document processor
Start-DocumentProcessor

# Check system status
Get-DoclingSystemStatus
```

### Example: Using Default Configuration

```powershell
# 1. Start the backend services
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem

# 2. Verify services are running
$status = Get-DoclingSystemStatus
if ($status.Backend.APIHealthy) {
    Write-Host "Backend API is running on port 8080" -ForegroundColor Green
}

# 3. Test the API
$response = Invoke-RestMethod -Uri "http://localhost:8080/api/health"
Write-Host "API Status: $($response.status)"
```

## Custom Domain and Ports

### Configuration for Custom Ports

```powershell
# Method 1: Using Start-All.ps1 with custom ports
.\Start-All.ps1 -ApiPort 9080 -WebPort 9081

# Method 2: Using module functions
Import-Module PSDocling
$script:DoclingSystem.APIPort = 9080
$script:DoclingSystem.WebPort = 9081
Initialize-DoclingSystem
Start-DoclingSystem
```

### Setting Up Custom Domain (Requires Admin)

```powershell
# 1. Run PowerShell as Administrator

# 2. Add URL ACL for custom domain
netsh http add urlacl url=http://myserver.local:9080/ user=Everyone

# 3. Update configuration
$config = Get-DoclingConfiguration
$config.Server.APIHost = 'myserver.local'
$config.Server.APIPort = 9080
Set-DoclingConfiguration -Config $config

# 4. Start services with URL ACL
.\Start-All.ps1 -EnsureUrlAcl -ApiPort 9080
```

### Example: Production Setup

```powershell
# Production configuration example
# Run as Administrator

# 1. Set up custom domain and ports
$customHost = "docling.company.com"
$customPort = 8443

# 2. Add URL ACL
netsh http add urlacl url=http://${customHost}:${customPort}/ user=Everyone

# 3. Configure firewall (if needed)
New-NetFirewallRule -DisplayName "PSDocling API" `
    -Direction Inbound `
    -LocalPort $customPort `
    -Protocol TCP `
    -Action Allow

# 4. Update configuration file
@{
    Server = @{
        APIHost = $customHost
        APIPort = $customPort
        Protocol = 'http'
        EnableCORS = 'true'
        AllowedOrigins = @("http://${customHost}:8081", "https://app.company.com")
    }
} | Export-PowerShellDataFile -Path ".\PSDocling.config.psd1"

# 5. Start services
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem
```

## API Endpoints

### Available Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/health` | GET | Health check |
| `/api/status` | GET | System status and queue info |
| `/api/documents` | GET | List all documents |
| `/api/upload` | POST | Upload document for processing |
| `/api/files` | GET | List processed files |
| `/api/download/{id}` | GET | Download processed file |

### Testing API Endpoints

```powershell
# Health check
curl http://localhost:8080/api/health

# System status
$status = Invoke-RestMethod -Uri "http://localhost:8080/api/status"
Write-Host "Queue Count: $($status.queueCount)"
Write-Host "Processing: $($status.processingCount)"

# List documents
$docs = Invoke-RestMethod -Uri "http://localhost:8080/api/documents"
$docs | Format-Table FileName, Status, Progress -AutoSize
```

## Health Monitoring

### Checking Service Health

```powershell
# Method 1: Using module function
$status = Get-DoclingSystemStatus
if ($status.Backend.APIHealthy) {
    Write-Host "API is healthy" -ForegroundColor Green
}
Write-Host "Queue Count: $($status.Backend.QueueCount)"
Write-Host "Processing: $($status.Backend.ProcessingCount)"

# Method 2: Direct API call
try {
    $response = Invoke-WebRequest -Uri "http://localhost:8080/api/health" -UseBasicParsing
    if ($response.StatusCode -eq 200) {
        Write-Host "Backend is responding" -ForegroundColor Green
    }
} catch {
    Write-Host "Backend is not responding" -ForegroundColor Red
}
```

### Monitoring Document Processing

```powershell
# Watch queue and processing status
while ($true) {
    Clear-Host
    $status = Get-DoclingSystemStatus
    Write-Host "=== PSDocling Backend Status ===" -ForegroundColor Cyan
    Write-Host "API Status: $(if($status.Backend.APIHealthy){'Connected'}else{'Disconnected'})"
    Write-Host "Queue: $($status.Backend.QueueCount) items"
    Write-Host "Processing: $($status.Backend.ProcessingCount) items"
    Write-Host "Completed: $($status.System.TotalDocumentsProcessed) items"
    Write-Host ""
    Write-Host "Press Ctrl+C to stop monitoring"
    Start-Sleep -Seconds 2
}
```

## Stopping Services

### Graceful Shutdown

```powershell
# Stop all services
.\Stop-All.ps1

# Or manually
Stop-Process -Name powershell -Force | Where-Object {
    $_.CommandLine -like "*Start-APIServer*" -or
    $_.CommandLine -like "*Start-DocumentProcessor*"
}
```

### Cleanup

```powershell
# Clear processing queue and status
Clear-PSDoclingSystem -Force

# Remove temporary files
Remove-Item "$env:TEMP\docling_*.json" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:TEMP\DoclingProcessor" -Recurse -Force -ErrorAction SilentlyContinue
```

## Troubleshooting

### Common Issues

1. **Port Already in Use**
   ```powershell
   # Check what's using the port
   netstat -ano | findstr :8080

   # Use different port
   .\Start-All.ps1 -ApiPort 9080
   ```

2. **Access Denied on Custom Domain**
   ```powershell
   # Run as Administrator and add URL ACL
   netsh http add urlacl url=http://+:8080/ user=Everyone
   ```

3. **Python Not Found (Simulation Mode)**
   ```powershell
   # Skip Python check for testing
   .\Start-All.ps1 -SkipPythonCheck
   ```

### Debug Information

Check debug files when troubleshooting:
- Python errors: `Get-Content "$env:TEMP\docling_error.txt"`
- Python output: `Get-Content "$env:TEMP\docling_output.txt"`
- Queue status: `Get-ChildItem "$env:TEMP\DoclingQueue" -Filter "*.queue" | Sort-Object CreationTime`