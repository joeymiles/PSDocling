# Backend Services Guide

## Overview
PSDocling runs two hidden PowerShell processes: the API server, which also serves the web UI, and the document processor. This guide covers starting, calling and stopping them.

## Table of Contents
- [Starting Backend Services](#starting-backend-services)
- [Default Configuration](#default-configuration)
- [Security Defaults](#security-defaults)
- [API Endpoints](#api-endpoints)
- [Health Monitoring](#health-monitoring)
- [Stopping Services](#stopping-services)
- [Troubleshooting](#troubleshooting)

## Starting Backend Services

### Prerequisites
- PowerShell 5.1 or higher
- Python 3.8+ with docling package (optional for simulation mode)
- No administrator rights are needed

### Quick Start

```powershell
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem               # headless: API + processor only
Start-DoclingSystem -UseWebView   # plus the app window (closing it stops everything)
```

For development from a checkout (uses the checkout's build, visible console):
```powershell
.\app\scripts\Start-All.ps1 -OpenBrowser
```

## Default Configuration

### Addresses and Paths
- **API server and UI**: `http://localhost:8080` (loopback only)
- **Document processor**: background process, no direct access
- **Data folder**: `%LOCALAPPDATA%\PSDocling` (set `PSDOCLING_HOME` to move it)
  - Queue folder: `data\queue` (one `.queue` file per job)
  - Status: `data\status.json`
  - Uploads: `data\uploads`, converted output: `data\output`
  - Logs: `logs\` (api, processor, launcher, window; capped with one rotation)
  - Runtime: `run\` (token, process ids, helper scripts)

### Another Port

```powershell
Start-DoclingSystem -Port 9080 -UseWebView
```

### Starting Individual Services

```powershell
# API server only (blocks the current session)
Start-APIServer -Port 8080

# Document processor only (blocks the current session)
Start-DocumentProcessor

# Check system status
Get-DoclingSystemStatus
```

## Security Defaults

PSDocling is a local desktop app:

- The API binds to `localhost` and rejects requests whose `Host` header is not `localhost` or `127.0.0.1`.
- The UI is served by the API itself, so there is no cross-origin (CORS) access for other web pages.
- Requests that change anything (POST) need the per-run token in the `X-PSDocling-Token` header. The token is created at each start and written to `run\token.txt`; it is removed on shutdown.
- Serving PSDocling to other machines is not supported.

```powershell
$token = Get-Content "$env:LOCALAPPDATA\PSDocling\run\token.txt"
$auth = @{ 'X-PSDocling-Token' = $token }
```

## API Endpoints

| Endpoint | Method | Token | Description |
|----------|--------|-------|-------------|
| `/` | GET | | Web UI |
| `/api/health` | GET | | Health check |
| `/api/status` | GET | | Queue counts |
| `/api/documents` | GET | | All documents and their status |
| `/api/files` | GET | | Processed files |
| `/api/error/{id}` | GET | | Error details for a document |
| `/api/result/{id}` | GET | | Converted result |
| `/api/download/{id}` | GET | | Result folder as ZIP |
| `/api/download-all` | GET | | All results as ZIP |
| `/api/app-info` | GET | | Version, data folder, shortcut state |
| `/api/upload` | POST | yes | Upload a document (base64 JSON) |
| `/api/start-conversion` | POST | yes | Queue an uploaded document |
| `/api/reprocess` | POST | yes | Convert again with new options |
| `/api/cancel/{id}` | POST | yes | Cancel processing |
| `/api/documents/{id}/reset` | POST | yes | Move a document back to results |
| `/api/shortcut` | POST | yes | `{"create": true}` adds the desktop shortcut |
| `/api/shutdown` | POST | yes | Stop the API, processor and window |
| `/api/uninstall` | POST | yes | Uninstall and delete data (the app confirms first) |

### Calling the API

```powershell
Invoke-RestMethod http://localhost:8080/api/health

$status = Invoke-RestMethod http://localhost:8080/api/status
"Queued: $($status.QueuedCount)  Processing: $($status.ProcessingCount)"

$docs = Invoke-RestMethod http://localhost:8080/api/documents
$docs | Format-Table fileName, status, progress -AutoSize

# Upload and convert (needs the token)
$token = Get-Content "$env:LOCALAPPDATA\PSDocling\run\token.txt"
$auth = @{ 'X-PSDocling-Token' = $token }
$body = @{ fileName = 'report.pdf'; dataBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes('C:\docs\report.pdf')) } | ConvertTo-Json
$upload = Invoke-RestMethod http://localhost:8080/api/upload -Method POST -Headers $auth -ContentType 'application/json' -Body $body
Invoke-RestMethod http://localhost:8080/api/start-conversion -Method POST -Headers $auth -ContentType 'application/json' -Body (@{ documentId = $upload.documentId } | ConvertTo-Json)
```

## Health Monitoring

```powershell
$status = Get-DoclingSystemStatus
if ($status.Backend.APIHealthy) { Write-Host "API is healthy" -ForegroundColor Green }
"Queue: $($status.Backend.QueueCount)  Processing: $($status.Backend.ProcessingCount)"
```

## Stopping Services

```powershell
Stop-DoclingSystem              # stop processes, keep queue and history
Stop-DoclingSystem -ClearQueue  # also clear queue and history
```

In the app, **Quit** does the same as `Stop-DoclingSystem`. Closing the window also stops everything. When started with a window, the API also stops itself after 3 minutes without any request from the UI.

### Cleanup

```powershell
# Clear queue and history (asks before deleting output unless -Force)
Clear-PSDoclingSystem
```

To remove PSDocling and all its data, use **Settings > Uninstall PSDocling** or `Uninstall-PSDocling.ps1` in the module folder.

## Troubleshooting

1. **Port already in use**
   ```powershell
   netstat -ano | findstr :8080
   Start-DoclingSystem -Port 9080 -UseWebView
   ```

2. **403 "Missing or invalid token"**: POST requests need `X-PSDocling-Token` from `run\token.txt`; the token changes at every start.

3. **Python not found (simulation mode)**
   ```powershell
   Initialize-DoclingSystem -SkipPythonCheck
   ```

### Debug Information
- Last failed start: `logs\last-launch-error.txt`
- API and launcher: `logs\api.log`, `logs\launcher.log`, `logs\window.log`
- Processor: `logs\processor.log`, `logs\processor-debug.log`, `logs\processor-errors.log`
- Last Python run: `run\docling_output.txt`, `run\docling_error.txt`
- Queue: `Get-ChildItem "$env:LOCALAPPDATA\PSDocling\data\queue" -Filter *.queue | Sort-Object CreationTime`
