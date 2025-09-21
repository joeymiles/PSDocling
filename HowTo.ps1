# PSDocling - PowerShell Document Processing System
# Comprehensive Usage Examples and How-To Guide

#region Installation Examples

Write-Host "=== INSTALLATION METHODS ===" -ForegroundColor Cyan

# Method 1: Install from PowerShell Gallery (Recommended)
Write-Host "`n1. Install from PowerShell Gallery:" -ForegroundColor Yellow
Write-Host @"
# Install for current user
Install-Module PSDocling -Scope CurrentUser

# Install for all users (requires admin)
Install-Module PSDocling -Scope AllUsers

# Import and start
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem -OpenBrowser
"@

# Method 2: Install from GitHub
Write-Host "`n2. Install from GitHub:" -ForegroundColor Yellow
Write-Host @"
# Clone repository
git clone https://github.com/joeymiles/PSDocling.git
cd PSDocling

# Install using provided script
.\Install-DoclingModule.ps1

# Or import directly
Import-Module .\PSDocling.psm1
"@

# Method 3: Manual Installation
Write-Host "`n3. Manual Installation:" -ForegroundColor Yellow
Write-Host @"
# Download and extract to PowerShell modules directory
# Windows PowerShell: `$env:USERPROFILE\Documents\WindowsPowerShell\Modules\PSDocling\
# PowerShell Core: `$env:USERPROFILE\Documents\PowerShell\Modules\PSDocling\
"@

#endregion

#region Quick Start Examples

Write-Host "`n=== QUICK START EXAMPLES ===" -ForegroundColor Cyan

# Basic startup
Write-Host "`n1. Basic Startup:" -ForegroundColor Yellow
Write-Host @"
Import-Module PSDocling
Initialize-DoclingSystem -GenerateFrontend
Start-DoclingSystem -OpenBrowser
"@

# Advanced startup with custom ports
Write-Host "`n2. Advanced Startup:" -ForegroundColor Yellow
Write-Host @"
# Custom ports and options
.\Start-All.ps1 -ApiPort 9080 -WebPort 9081 -GenerateFrontend -OpenBrowser

# Skip Python check (simulation mode)
.\Start-All.ps1 -SkipPythonCheck

# Ensure URL ACL for HTTP listener (requires admin)
.\Start-All.ps1 -EnsureUrlAcl
"@

#endregion

#region Document Processing Examples

Write-Host "`n=== DOCUMENT PROCESSING EXAMPLES ===" -ForegroundColor Cyan

Write-Host "`n1. Basic Document Processing:" -ForegroundColor Yellow
Write-Host @"
# Add single document
Add-DocumentToQueue -Path "C:\Documents\sample.pdf"

# Add multiple documents
Add-DocumentToQueue -Path "C:\Documents\presentation.pptx"
Add-DocumentToQueue -Path "C:\Documents\spreadsheet.xlsx"
Add-DocumentToQueue -Path "C:\Documents\document.docx"

# Check processing status
Get-DoclingSystemStatus
"@

Write-Host "`n2. Queue Management:" -ForegroundColor Yellow
Write-Host @"
# View current queue
Get-QueueItems

# Get next item for processing
Get-NextQueueItem

# Check specific item status
Get-ProcessingStatus

# Get system status and statistics
Get-DoclingSystemStatus
"@

#endregion

#region REST API Examples

Write-Host "`n=== REST API USAGE EXAMPLES ===" -ForegroundColor Cyan

Write-Host "`n1. File Upload via API:" -ForegroundColor Yellow
Write-Host @"
# Upload a single file
`$response = Invoke-RestMethod -Uri "http://localhost:8080/upload" -Method Post -InFile "C:\Documents\sample.pdf"
Write-Host "Upload ID: `$(`$response.id)"

# Upload with custom filename
`$body = @{
    file = Get-Item "C:\Documents\report.pdf"
    filename = "monthly-report.pdf"
}
`$response = Invoke-RestMethod -Uri "http://localhost:8080/upload" -Method Post -Body `$body

# Upload multiple files
`$files = @("document1.pdf", "document2.docx", "presentation.pptx")
foreach (`$file in `$files) {
    `$response = Invoke-RestMethod -Uri "http://localhost:8080/upload" -Method Post -InFile `$file
    Write-Host "Uploaded `$file with ID: `$(`$response.id)"
}
"@

Write-Host "`n2. Status and Queue Management via API:" -ForegroundColor Yellow
Write-Host @"
# Get system status
`$status = Invoke-RestMethod -Uri "http://localhost:8080/status"
Write-Host "Queue length: `$(`$status.queueLength)"
Write-Host "Processing: `$(`$status.processing)"

# Get current queue
`$queue = Invoke-RestMethod -Uri "http://localhost:8080/queue"
foreach (`$item in `$queue) {
    Write-Host "`$(`$item.FileName) - Status: `$(`$item.Status)"
}

# Get processing statistics
`$stats = Invoke-RestMethod -Uri "http://localhost:8080/stats"
Write-Host "Total processed: `$(`$stats.totalProcessed)"
Write-Host "Success rate: `$(`$stats.successRate)%"
"@

Write-Host "`n3. Download Processed Files via API:" -ForegroundColor Yellow
Write-Host @"
# List available processed files
`$files = Invoke-RestMethod -Uri "http://localhost:8080/files"
foreach (`$file in `$files) {
    Write-Host "Available: `$(`$file.name) (Size: `$(`$file.size) bytes)"
}

# Download specific processed file
`$fileId = "12345"
Invoke-RestMethod -Uri "http://localhost:8080/download/`$fileId" -OutFile "processed-document.md"

# Download with original filename preservation
`$downloadInfo = Invoke-RestMethod -Uri "http://localhost:8080/file-info/`$fileId"
`$originalName = `$downloadInfo.originalName -replace '\.[^.]+$', '.md'
Invoke-RestMethod -Uri "http://localhost:8080/download/`$fileId" -OutFile `$originalName
"@

Write-Host "`n4. Advanced API Usage:" -ForegroundColor Yellow
Write-Host @"
# Check if specific file was processed
`$fileName = "important-document.pdf"
`$processedFiles = Invoke-RestMethod -Uri "http://localhost:8080/files"
`$isProcessed = `$processedFiles | Where-Object { `$_.originalName -eq `$fileName }

if (`$isProcessed) {
    Write-Host "`$fileName has been processed"
    # Download the processed version
    Invoke-RestMethod -Uri "http://localhost:8080/download/`$(`$isProcessed.id)" -OutFile "processed-`$fileName.md"
} else {
    Write-Host "`$fileName not found in processed files"
}

# Bulk download all processed files
`$processedFiles = Invoke-RestMethod -Uri "http://localhost:8080/files"
New-Item -Path ".\Downloads" -ItemType Directory -Force
foreach (`$file in `$processedFiles) {
    `$outputPath = ".\Downloads\`$(`$file.name)"
    Invoke-RestMethod -Uri "http://localhost:8080/download/`$(`$file.id)" -OutFile `$outputPath
    Write-Host "Downloaded: `$outputPath"
}
"@

#endregion

#region Monitoring and Troubleshooting

Write-Host "`n=== MONITORING AND TROUBLESHOOTING ===" -ForegroundColor Cyan

Write-Host "`n1. System Health Checks:" -ForegroundColor Yellow
Write-Host @"
# Check Python and Docling availability
Get-PythonStatus

# Comprehensive system status
`$status = Get-DoclingSystemStatus
Write-Host "API Server: `$(`$status.ApiRunning)"
Write-Host "Web Server: `$(`$status.WebRunning)"
Write-Host "Processor: `$(`$status.ProcessorRunning)"
Write-Host "Queue Length: `$(`$status.QueueLength)"

# Test API connectivity
try {
    `$response = Invoke-RestMethod -Uri "http://localhost:8080/health" -TimeoutSec 5
    Write-Host "API Server: Online" -ForegroundColor Green
} catch {
    Write-Host "API Server: Offline" -ForegroundColor Red
}
"@

Write-Host "`n2. Log Monitoring:" -ForegroundColor Yellow
Write-Host @"
# Monitor queue file changes (manual method)
Get-Content "`$env:TEMP\docling_queue.json" | ConvertFrom-Json

# Monitor status file
Get-Content "`$env:TEMP\docling_status.json" | ConvertFrom-Json

# Check temp directory for processing files
Get-ChildItem "`$env:TEMP\DoclingProcessor" | Sort-Object LastWriteTime -Descending
"@

#endregion

#region Production Deployment Examples

Write-Host "`n=== PRODUCTION DEPLOYMENT EXAMPLES ===" -ForegroundColor Cyan

Write-Host "`n1. Service-Style Deployment:" -ForegroundColor Yellow
Write-Host @"
# Create a startup script for production
`$startupScript = @'
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem
'@
`$startupScript | Out-File -FilePath "Start-PSDocling-Service.ps1"

# Run as background job
Start-Job -ScriptBlock {
    Import-Module PSDocling
    Initialize-DoclingSystem
    Start-DoclingSystem
}
"@

Write-Host "`n2. Scheduled Task Deployment:" -ForegroundColor Yellow
Write-Host @"
# Register as Windows scheduled task (requires admin)
`$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\PSDocling\Start-PSDocling-Service.ps1"
`$trigger = New-ScheduledTaskTrigger -AtStartup
`$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount
Register-ScheduledTask -TaskName "PSDocling" -Action `$action -Trigger `$trigger -Principal `$principal
"@

#endregion

#region Integration Examples

Write-Host "`n=== INTEGRATION EXAMPLES ===" -ForegroundColor Cyan

Write-Host "`n1. Automated Document Processing Pipeline:" -ForegroundColor Yellow
Write-Host @"
# Watch folder for new documents and auto-process
`$watcher = New-Object System.IO.FileSystemWatcher
`$watcher.Path = "C:\IncomingDocuments"
`$watcher.Filter = "*.pdf"
`$watcher.EnableRaisingEvents = `$true

Register-ObjectEvent -InputObject `$watcher -EventName Created -Action {
    `$filePath = `$Event.SourceEventArgs.FullPath
    Write-Host "New file detected: `$filePath"
    Add-DocumentToQueue -Path `$filePath
}
"@

Write-Host "`n2. Integration with Office 365:" -ForegroundColor Yellow
Write-Host @"
# Process SharePoint documents (requires SharePoint PowerShell module)
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite"
`$files = Get-PnPFolderItem -FolderSiteRelativeUrl "Shared Documents"

foreach (`$file in `$files | Where-Object { `$_.Name -match '\.(pdf|docx|xlsx|pptx)`$' }) {
    `$tempPath = "`$env:TEMP\`$(`$file.Name)"
    Get-PnPFile -Url `$file.ServerRelativeUrl -Path `$env:TEMP -Filename `$file.Name -AsFile
    Add-DocumentToQueue -Path `$tempPath
}
"@

#endregion

#region Cleanup and Maintenance

Write-Host "`n=== CLEANUP AND MAINTENANCE ===" -ForegroundColor Cyan

Write-Host "`n1. System Cleanup:" -ForegroundColor Yellow
Write-Host @"
# Stop all services
.\Stop-All.ps1

# Clean temporary files
Remove-Item "`$env:TEMP\docling_*" -Force -ErrorAction SilentlyContinue
Remove-Item "`$env:TEMP\DoclingProcessor" -Recurse -Force -ErrorAction SilentlyContinue

# Clean old processed files (older than 7 days)
Get-ChildItem ".\ProcessedDocuments" | Where-Object { `$_.LastWriteTime -lt (Get-Date).AddDays(-7) } | Remove-Item
"@

Write-Host "`n2. Module Maintenance:" -ForegroundColor Yellow
Write-Host @"
# Update from PowerShell Gallery
Update-Module PSDocling

# Uninstall module
Uninstall-Module PSDocling

# Or use provided uninstall script
.\Uninstall-DoclingModule.ps1 -Scope CurrentUser -Force
"@

#endregion

Write-Host "`n=== END OF EXAMPLES ===" -ForegroundColor Green
Write-Host "For more information, visit: https://github.com/joeymiles/PSDocling" -ForegroundColor Cyan

# Uncomment the lines below to run actual examples
<#
# Basic example - uncomment to run
Import-Module PSDocling
Initialize-DoclingSystem -GenerateFrontend
Start-DoclingSystem -OpenBrowser

# Add a sample document (update path as needed)
# Add-DocumentToQueue -Path "C:\Users\$env:USERNAME\Downloads\sample.pdf"

# Check system status
Get-DoclingSystemStatus
#>