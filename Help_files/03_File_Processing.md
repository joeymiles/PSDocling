# File Processing Guide

## Overview
This guide covers the complete file processing workflow in PSDocling, from uploading documents to retrieving processed results. Learn how to add files, monitor progress, and access converted documents.

## Table of Contents
- [Supported File Formats](#supported-file-formats)
- [Adding Files for Processing](#adding-files-for-processing)
- [Monitoring Progress](#monitoring-progress)
- [Accessing Processed Files](#accessing-processed-files)
- [Processing Options](#processing-options)
- [Advanced Processing](#advanced-processing)
- [Troubleshooting](#troubleshooting)

## Supported File Formats

### Document Formats
- **PDF** (.pdf) - Full support with text, images, tables
- **Microsoft Word** (.docx) - Complete document structure
- **Microsoft Excel** (.xlsx) - Tables and data preservation
- **Microsoft PowerPoint** (.pptx) - Slides and content extraction
- **Markdown** (.md) - Plain text with formatting
- **HTML** (.html, .xhtml) - Web content conversion
- **CSV** (.csv) - Structured data files

### Image Formats
- **PNG** (.png)
- **JPEG** (.jpg, .jpeg)
- **TIFF** (.tiff)
- **BMP** (.bmp)
- **WebP** (.webp)

### Output Formats
- **Markdown** - Human-readable with preserved formatting
- **JSON** - Structured data for programmatic access
- **HTML** - Web-ready with embedded styling
- **DocTags** - XML format for data extraction

## Adding Files for Processing

### Method 1: PowerShell Command Line

```powershell
# Add single file
Add-DocumentToQueue -Path "C:\Documents\report.pdf"

# Add multiple files
$files = Get-ChildItem "C:\Documents" -Filter "*.pdf"
foreach ($file in $files) {
    Add-DocumentToQueue -Path $file.FullName
}

# Add with specific output format
Add-DocumentToQueue -Path "C:\Documents\presentation.pptx" -OutputFormat "json"
```

### Method 2: REST API

```powershell
# Upload via API
$file = Get-Item "C:\Documents\sample.pdf"
$response = Invoke-RestMethod -Uri "http://localhost:8080/api/upload" `
    -Method Post `
    -InFile $file.FullName

Write-Host "Document ID: $($response.id)"
Write-Host "Status: $($response.status)"
```

### Method 3: Web Interface

```text
1. Navigate to http://localhost:8081
2. Drag and drop files onto upload area
   OR
   Click "Choose Files" and browse
3. Select processing options
4. Click "Process Documents"
```

### Method 4: Batch Processing Script

```powershell
# Create a batch processing script
$documentsPath = "C:\BatchDocuments"
$outputFormat = "markdown"

# Start services
Import-Module PSDocling
Initialize-DoclingSystem
Start-DoclingSystem

# Process all PDFs in folder
Get-ChildItem $documentsPath -Filter "*.pdf" | ForEach-Object {
    Write-Host "Adding $($_.Name) to queue..."
    Add-DocumentToQueue -Path $_.FullName
}

# Monitor until complete
do {
    Start-Sleep -Seconds 5
    $status = Get-DoclingSystemStatus
    Write-Host "Processing: $($status.Backend.ProcessingCount) | Queue: $($status.Backend.QueueCount)"
} while ($status.Backend.QueueCount -gt 0 -or $status.Backend.ProcessingCount -gt 0)

Write-Host "Batch processing complete!"
```

## Monitoring Progress

### Real-Time Status Monitoring

```powershell
# Check overall system status
Get-DoclingSystemStatus

# Get detailed processing status
$status = Get-ProcessingStatus
foreach ($id in $status.Keys) {
    $item = $status[$id]
    Write-Host "$($item.FileName): $($item.Status) - $($item.Progress)%"
}

# Monitor specific document
$documentId = "305d7273-145f-4614-80ef-9933cfec0506"
$docStatus = (Get-ProcessingStatus)[$documentId]
Write-Host "File: $($docStatus.FileName)"
Write-Host "Status: $($docStatus.Status)"
Write-Host "Progress: $($docStatus.Progress)%"
```

### Progress Tracking Script

```powershell
# Continuous monitoring script
param($DocumentId)

while ($true) {
    Clear-Host
    $status = Get-ProcessingStatus

    if ($DocumentId) {
        # Monitor specific document
        $doc = $status[$DocumentId]
        if ($doc) {
            Write-Host "=== Document Progress ===" -ForegroundColor Cyan
            Write-Host "File: $($doc.FileName)"
            Write-Host "Status: $($doc.Status)"
            Write-Host "Progress: $($doc.Progress)%"

            if ($doc.Status -eq "Completed") {
                Write-Host "`nProcessing complete!" -ForegroundColor Green
                Write-Host "Output: $($doc.OutputFile)"
                break
            }
            elseif ($doc.Status -eq "Failed") {
                Write-Host "`nProcessing failed!" -ForegroundColor Red
                Write-Host "Error: $($doc.Error)"
                break
            }
        }
    }
    else {
        # Monitor all documents
        Write-Host "=== Processing Queue ===" -ForegroundColor Cyan
        foreach ($id in $status.Keys) {
            $doc = $status[$id]
            $color = switch($doc.Status) {
                "Completed" { "Green" }
                "Processing" { "Yellow" }
                "Failed" { "Red" }
                default { "White" }
            }
            Write-Host "$($doc.FileName): $($doc.Status) ($($doc.Progress)%)" -ForegroundColor $color
        }
    }

    Start-Sleep -Seconds 2
}
```

### Processing States

| Status | Description |
|--------|-------------|
| **Queued** | Document waiting to be processed |
| **Processing** | Currently being converted |
| **Completed** | Successfully processed |
| **Failed** | Error during processing |

## Accessing Processed Files

### Finding Processed Documents

```powershell
# List all processed files
$processedDir = ".\ProcessedDocuments"
Get-ChildItem $processedDir -Recurse -File |
    Select-Object Name, Directory, Length, LastWriteTime |
    Format-Table -AutoSize

# Find specific document by original name
$originalName = "report.pdf"
$processed = Get-ChildItem $processedDir -Recurse -Filter "*.md" |
    Where-Object { $_.Name -like "$($originalName -replace '\.pdf$','*')" }

Write-Host "Found processed file: $($processed.FullName)"
```

### Copying Processed Files

```powershell
# Copy single processed file
$documentId = "305d7273-145f-4614-80ef-9933cfec0506" # example ID
$source = ".\ProcessedDocuments\$documentId\*.md"
$destination = "C:\Output\"

Copy-Item $source $destination -Force
Write-Host "Copied to: $destination"

# Copy all processed files from today
$today = (Get-Date).Date
$processed = Get-ChildItem ".\ProcessedDocuments" -Recurse -File |
    Where-Object { $_.CreationTime.Date -eq $today }

foreach ($file in $processed) {
    Copy-Item $file.FullName "C:\TodaysDocuments\" -Force
    Write-Host "Copied: $($file.Name)"
}
```

### Downloading via API

```powershell
# Get list of available files
$files = Invoke-RestMethod -Uri "http://localhost:8080/api/files"
$files | Format-Table name, size, created -AutoSize

# Download specific file
$fileId = $files[0].id
$outputPath = "C:\Downloads\$($files[0].name)"
Invoke-RestMethod -Uri "http://localhost:8080/api/download/$fileId" `
    -OutFile $outputPath

Write-Host "Downloaded to: $outputPath"
```

### Organized Output Structure

```powershell
# Organize processed files by type
$processedDir = ".\ProcessedDocuments"
$organizedDir = "C:\OrganizedOutput"

# Create folders by output type
$types = @("markdown", "json", "html", "xml")
foreach ($type in $types) {
    New-Item -Path "$organizedDir\$type" -ItemType Directory -Force
}

# Move files to organized structure
Get-ChildItem $processedDir -Recurse -File | ForEach-Object {
    $extension = $_.Extension.TrimStart('.')
    $targetDir = "$organizedDir\$extension"

    if (Test-Path $targetDir) {
        Copy-Item $_.FullName "$targetDir\$($_.Name)" -Force
        Write-Host "Organized: $($_.Name) -> $extension folder"
    }
}
```

## Processing Options

### Basic Options

```powershell
# Markdown output (default)
Add-DocumentToQueue -Path "document.pdf"

# JSON output for data processing
Add-DocumentToQueue -Path "document.pdf" -OutputFormat "json"

# HTML output for web display
Add-DocumentToQueue -Path "document.pdf" -OutputFormat "html"

# DocTags XML format
Add-DocumentToQueue -Path "document.pdf" -OutputFormat "doctags"
```

### Enrichment Options

```powershell
# Enable code understanding
$params = @{
    Path = "technical-manual.pdf"
    EnableCodeEnrichment = $true
}
Add-DocumentToQueue @params

# Enable formula detection
$params = @{
    Path = "scientific-paper.pdf"
    EnableFormulaEnrichment = $true
}
Add-DocumentToQueue @params

# Enable picture analysis
$params = @{
    Path = "illustrated-guide.pdf"
    EnablePictureClassification = $true
    EnablePictureDescription = $true
}
Add-DocumentToQueue @params

# All enrichments
$params = @{
    Path = "complex-document.pdf"
    EnableCodeEnrichment = $true
    EnableFormulaEnrichment = $true
    EnablePictureClassification = $true
    EnablePictureDescription = $true
}
Add-DocumentToQueue @params
```

### Image Handling

```powershell
# Embed images in output
$params = @{
    Path = "document-with-images.pdf"
    EmbedImages = $true
}
Add-DocumentToQueue @params

# Extract images as separate files (default)
$params = @{
    Path = "document-with-images.pdf"
    EmbedImages = $false
}
Add-DocumentToQueue @params
```

## Advanced Processing

### Chunking for Large Documents

```powershell
# Enable chunking for large files
Invoke-DoclingHybridChunking -FilePath "large-document.pdf" `
    -MaxChunkSize 1500 `
    -MinChunkSize 300 `
    -OutputFormat "markdown"

# Process with semantic chunking
$result = Invoke-DoclingHybridChunking -FilePath "technical-manual.pdf" `
    -EnableSemanticChunking `
    -PreserveSentences `
    -OutputFormat "json"

Write-Host "Created $($result.chunks.Count) chunks"
```

### Custom Processing Pipeline

```powershell
# Create custom processing function
function Process-DocumentWithRetry {
    param(
        [string]$FilePath,
        [string]$OutputFormat = "markdown",
        [int]$MaxRetries = 3
    )

    $attempt = 0
    $success = $false

    while ($attempt -lt $MaxRetries -and -not $success) {
        $attempt++
        Write-Host "Processing attempt $attempt of $MaxRetries..."

        # Add to queue
        $id = (Add-DocumentToQueue -Path $FilePath).Id

        # Wait for completion
        $timeout = New-TimeSpan -Minutes 10
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        while ($stopwatch.Elapsed -lt $timeout) {
            $status = (Get-ProcessingStatus)[$id]

            if ($status.Status -eq "Completed") {
                $success = $true
                Write-Host "Successfully processed!" -ForegroundColor Green
                return $status.OutputFile
                break
            }
            elseif ($status.Status -eq "Failed") {
                Write-Host "Attempt $attempt failed" -ForegroundColor Yellow
                break
            }

            Start-Sleep -Seconds 5
        }
    }

    if (-not $success) {
        throw "Failed to process document after $MaxRetries attempts"
    }
}

# Use the custom function
$outputFile = Process-DocumentWithRetry -FilePath "important.pdf"
Write-Host "Output saved to: $outputFile"
```

### Parallel Batch Processing

```powershell
# Process multiple documents in parallel queues
$documents = Get-ChildItem "C:\Documents" -Filter "*.pdf"

# Add all to queue first
$documentIds = @()
foreach ($doc in $documents) {
    $result = Add-DocumentToQueue -Path $doc.FullName
    $documentIds += @{
        Id = $result.Id
        Name = $doc.Name
        Status = "Queued"
    }
}

# Monitor all documents
while ($documentIds | Where-Object { $_.Status -ne "Completed" }) {
    foreach ($doc in $documentIds) {
        if ($doc.Status -ne "Completed") {
            $status = (Get-ProcessingStatus)[$doc.Id]
            $doc.Status = $status.Status
            Write-Host "$($doc.Name): $($status.Status) - $($status.Progress)%"
        }
    }
    Start-Sleep -Seconds 3
    Clear-Host
}

Write-Host "All documents processed!" -ForegroundColor Green
```

## Troubleshooting

### Common Processing Issues

1. **Document Stuck in Queue**
   ```powershell
   # Check if processor is running
   Get-Process | Where-Object {$_.CommandLine -like "*Start-DocumentProcessor*"}

   # Restart processor
   .\Stop-All.ps1
   .\Start-All.ps1
   ```

2. **Processing Failed**
   ```powershell
   # Check error logs
   Get-Content "$env:TEMP\docling_error.txt" -Tail 50

   # Get specific document error
   $status = Get-ProcessingStatus
   $failed = $status.Values | Where-Object { $_.Status -eq "Failed" }
   $failed | Select-Object FileName, Error
   ```

3. **Output File Missing**
   ```powershell
   # Verify output location
   $status = Get-ProcessingStatus
   $completed = $status.Values | Where-Object { $_.Status -eq "Completed" }
   foreach ($doc in $completed) {
       if (-not (Test-Path $doc.OutputFile)) {
           Write-Host "Missing: $($doc.OutputFile)" -ForegroundColor Red
       }
   }
   ```

### Performance Optimization

```powershell
# Check processing times
$status = Get-ProcessingStatus
$completed = $status.Values | Where-Object { $_.Status -eq "Completed" }
$completed | ForEach-Object {
    [PSCustomObject]@{
        File = $_.FileName
        Size = [math]::Round($_.FileSize / 1MB, 2)
        Time = [math]::Round($_.ElapsedTime / 1000, 2)
        Speed = [math]::Round($_.FileSize / $_.ElapsedTime, 2)
    }
} | Format-Table -AutoSize
```

### Recovery Scripts

```powershell
# Clear stuck queue
Clear-PSDoclingSystem -Force

# Reprocess failed documents
$status = Get-ProcessingStatus
$failed = $status.Values | Where-Object { $_.Status -eq "Failed" }
foreach ($doc in $failed) {
    Write-Host "Reprocessing: $($doc.FileName)"
    Add-DocumentToQueue -Path $doc.FilePath
}
```