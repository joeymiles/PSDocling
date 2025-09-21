# Docling Document Processing System
# Version: 2.1.1

$script:DoclingSystem = @{
    Version = "2.1.1"
    TempDirectory = "$env:TEMP\DoclingProcessor"
    OutputDirectory = ".\ProcessedDocuments"
    APIPort = 8080
    WebPort = 8081
    QueueFile = "$env:TEMP\docling_queue.json"
    StatusFile = "$env:TEMP\docling_status.json"
    PythonAvailable = $false
    ProcessingStatus = @{}
}

# Auto-detect Python on module load
try {
    $pythonCmd = Get-Command python -ErrorAction Stop
    if ($pythonCmd) {
        $script:DoclingSystem.PythonAvailable = $true
        Write-Host "Python auto-detected" -ForegroundColor Green
    }
} catch {
    Write-Host "Python not found during auto-detection" -ForegroundColor Yellow
}

# File-based queue management
function Get-QueueItems {
    if (Test-Path $script:DoclingSystem.QueueFile) {
        try {
            $content = Get-Content $script:DoclingSystem.QueueFile -Raw
            if ($content.Trim() -eq "[]") {
                return @()
            }
            # Force array conversion in PowerShell 5.1
            $items = @($content | ConvertFrom-Json)
            return $items
        } catch {
            return @()
        }
    }
    return @()
}

function Set-QueueItems {
    param([array]$Items = @())
    # Ensure we always store as a JSON array, even for single items
    if ($Items.Count -eq 0) {
        "[]" | Set-Content $script:DoclingSystem.QueueFile -Encoding UTF8
    } elseif ($Items.Count -eq 1) {
        "[" + ($Items[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $script:DoclingSystem.QueueFile -Encoding UTF8
    } else {
        $Items | ConvertTo-Json -Depth 10 | Set-Content $script:DoclingSystem.QueueFile -Encoding UTF8
    }
}

function Add-QueueItem {
    param($Item)
    $queue = Get-QueueItems
    # Force array addition
    $newQueue = @($queue) + @($Item)
    Set-QueueItems $newQueue
}

function Get-NextQueueItem {
    $queue = Get-QueueItems
    if ($queue.Count -gt 0) {
        $item = $queue[0]
        $remaining = if ($queue.Count -gt 1) { $queue[1..($queue.Count-1)] } else { @() }
        Set-QueueItems $remaining
        return $item
    }
    return $null
}

# File-based status management for cross-process sharing
function Get-ProcessingStatus {
    if (Test-Path $script:DoclingSystem.StatusFile) {
        try {
            $content = Get-Content $script:DoclingSystem.StatusFile -Raw
            $jsonObj = $content | ConvertFrom-Json

            # Convert PSCustomObject to hashtable manually
            $hashtable = @{}
            $jsonObj.PSObject.Properties | ForEach-Object {
                $hashtable[$_.Name] = $_.Value
            }
            return $hashtable
        } catch {
            return @{}
        }
    }
    return @{}
}

function Set-ProcessingStatus {
    param([hashtable]$Status)
    $Status | ConvertTo-Json -Depth 10 | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8
}

function Update-ItemStatus {
    param($Id, $Updates)
    $status = Get-ProcessingStatus

    # Convert existing item to hashtable if it's a PSObject
    if ($status[$Id]) {
        if ($status[$Id] -is [PSCustomObject]) {
            $itemHash = @{}
            $status[$Id].PSObject.Properties | ForEach-Object {
                $itemHash[$_.Name] = $_.Value
            }
            $status[$Id] = $itemHash
        }
    } else {
        $status[$Id] = @{}
    }

    # Apply updates
    foreach ($key in $Updates.Keys) {
        $status[$Id][$key] = $Updates[$key]
    }

    Set-ProcessingStatus $status
    # Also update local cache
    $script:DoclingSystem.ProcessingStatus[$Id] = $status[$Id]
}

function Initialize-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$SkipPythonCheck,
        [switch]$GenerateFrontend
    )

    Write-Host "Initializing PS Docling System v$($script:DoclingSystem.Version)" -ForegroundColor Cyan

    # Create directories
    @($script:DoclingSystem.TempDirectory, $script:DoclingSystem.OutputDirectory) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            Write-Host "Created directory: $_" -ForegroundColor Green
        }
    }

    # Initialize queue
    Set-QueueItems @()
    @{} | ConvertTo-Json | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8

    # Check Python
    if (-not $SkipPythonCheck) {
        try {
            $version = & python --version 2>&1
            if ($version -match "Python") {
                Write-Host "Python found: $version" -ForegroundColor Green
                $script:DoclingSystem.PythonAvailable = $true

                # Check Docling
                $pipOutput = & python -m pip show docling 2>&1
                if (-not ($pipOutput -match "Name: docling")) {
                    Write-Host "Installing Docling..." -ForegroundColor Yellow
                    & python -m pip install docling --quiet
                }
                Write-Host "Docling ready" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Python not found - using simulation mode"
        }
    }

    if ($GenerateFrontend) {
        New-FrontendFiles
    }

    Write-Host "System initialized" -ForegroundColor Green
}

function Add-DocumentToQueue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]$Path,
        [string]$ExportFormat = 'markdown'
    )

    process {
        foreach ($filePath in $Path) {
            if (Test-Path $filePath) {
                $fileInfo = Get-Item $filePath
                $supportedFormats = @('.pdf', '.docx', '.xlsx', '.pptx', '.md', '.html', '.xhtml', '.csv', '.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp', '.webp')

                if ($fileInfo.Extension -notin $supportedFormats) {
                    Write-Warning "Unsupported format: $($fileInfo.Extension)"
                    continue
                }

                $item = @{
                    Id = [guid]::NewGuid().ToString()
                    FilePath = $fileInfo.FullName
                    FileName = $fileInfo.Name
                    ExportFormat = $ExportFormat
                    Status = 'Ready'
                    UploadedTime = Get-Date
                }

                # Don't add to processing queue yet - just store status
                Update-ItemStatus $item.Id $item

                Write-Host "Queued: $($fileInfo.Name) (ID: $($item.Id))" -ForegroundColor Green
                Write-Output $item.Id
            }
        }
    }
}

function Start-DocumentConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DocumentId,
        [string]$ExportFormat
    )

    $allStatus = Get-ProcessingStatus
    $documentStatus = $allStatus[$DocumentId]

    if (-not $documentStatus) {
        Write-Error "Document not found: $DocumentId"
        return $false
    }

    if ($documentStatus.Status -ne 'Ready') {
        Write-Warning "Document $DocumentId is not in Ready status (current: $($documentStatus.Status))"
        return $false
    }

    # Update export format if provided
    if ($ExportFormat) {
        $documentStatus.ExportFormat = $ExportFormat
    }

    # Create queue item for processing
    $queueItem = @{
        Id = $DocumentId
        FilePath = $documentStatus.FilePath
        FileName = $documentStatus.FileName
        ExportFormat = $documentStatus.ExportFormat
        Status = 'Queued'
        QueuedTime = Get-Date
        UploadedTime = $documentStatus.UploadedTime
    }

    # Add to processing queue and update status
    Add-QueueItem $queueItem
    Update-ItemStatus $DocumentId @{
        Status = 'Queued'
        QueuedTime = Get-Date
        ExportFormat = $documentStatus.ExportFormat
    }

    Write-Host "Started conversion for: $($documentStatus.FileName) (ID: $DocumentId)" -ForegroundColor Green
    return $true
}

function Start-DocumentProcessor {
    [CmdletBinding()]
    param()

    Write-Host "Document processor started" -ForegroundColor Green

    while ($true) {
        $item = Get-NextQueueItem
        if ($item) {
            Write-Host "Processing: $($item.FileName)" -ForegroundColor Yellow

            # Update status
            Update-ItemStatus $item.Id @{
                Status = 'Processing'
                StartTime = Get-Date
            }

            # Create output directory
            $outputDir = Join-Path $script:DoclingSystem.OutputDirectory $item.Id
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)

            # Determine file extension based on export format
            $extension = switch ($item.ExportFormat) {
                'markdown' { '.md' }
                'html' { '.html' }
                'json' { '.json' }
                'text' { '.txt' }
                'doctags' { '.xml' }
                default { '.md' }
            }

            $outputFile = Join-Path $outputDir "$baseName$extension"

            $success = $false
            try {
                if ($script:DoclingSystem.PythonAvailable) {
                    # Create Python conversion script
                    $pyScript = @"
import sys
import json
from pathlib import Path

try:
    from docling.document_converter import DocumentConverter

    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    export_format = sys.argv[3] if len(sys.argv) > 3 else 'markdown'

    converter = DocumentConverter()
    result = converter.convert(str(src))

    # Export based on format
    if export_format == 'markdown':
        content = result.document.export_to_markdown()
    elif export_format == 'html':
        # HTML export may not be directly available, try markdown as fallback
        try:
            content = result.document.export_to_html()
        except AttributeError:
            content = result.document.export_to_markdown()
    elif export_format == 'json':
        import json
        content = json.dumps(result.document.export_to_dict(), indent=2, ensure_ascii=False)
    elif export_format == 'text':
        # Text export may not be directly available, extract from document
        try:
            content = result.document.export_to_text()
        except AttributeError:
            # Fallback: extract text from markdown
            md_content = result.document.export_to_markdown()
            # Simple markdown to text conversion (remove common markdown syntax)
            import re
            content = re.sub(r'[#*`_\[\]()]', '', md_content)
            content = re.sub(r'!\[.*?\]\(.*?\)', '', content)  # Remove images
            content = re.sub(r'\[.*?\]\(.*?\)', '', content)   # Remove links
            content = re.sub(r'\n+', '\n', content).strip()    # Clean up whitespace
    elif export_format == 'doctags':
        # DocTags export may not be directly available
        try:
            content = result.document.export_to_doctags()
        except AttributeError:
            # Fallback: create a simple XML representation
            import json
            doc_dict = result.document.export_to_dict()
            content = f'<?xml version="1.0" encoding="UTF-8"?>\n<document>\n<content><![CDATA[{json.dumps(doc_dict, indent=2)}]]></content>\n</document>'
    else:
        raise ValueError(f'Unsupported export format: {export_format}')

    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(content, encoding='utf-8')

    print(json.dumps({'success': True, 'format': export_format, 'output_file': str(dst)}))

except Exception as e:
    print(json.dumps({'success': False, 'error': str(e)}))
    sys.exit(1)
"@

                    $tempPy = Join-Path $env:TEMP "docling_$([guid]::NewGuid().ToString('N')[0..7] -join '').py"
                    $pyScript | Set-Content $tempPy -Encoding UTF8

                    try {
                        # Start Python process with timeout (10 minutes max)
                        # Use single argument string to properly handle spaces in filenames
                        $exportFormat = if ($item.ExportFormat) { $item.ExportFormat } else { 'markdown' }
                        $arguments = "`"$tempPy`" `"$($item.FilePath)`" `"$outputFile`" `"$exportFormat`""
                        $process = Start-Process python -ArgumentList $arguments -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\docling_output.txt" -RedirectStandardError "$env:TEMP\docling_error.txt"

                        # Wait up to 10 minutes (600 seconds) for complex PDFs
                        $finished = $process.WaitForExit(600000)

                        if ($finished) {
                            $stdout = Get-Content "$env:TEMP\docling_output.txt" -Raw -ErrorAction SilentlyContinue
                            $stderr = Get-Content "$env:TEMP\docling_error.txt" -Raw -ErrorAction SilentlyContinue

                            # Check if the output file was actually created and has content
                            $outputExists = (Test-Path $outputFile) -and ((Get-Item $outputFile).Length -gt 0)

                            # Also check Python JSON output for success
                            $pythonSuccess = $false
                            if ($stdout) {
                                try {
                                    $jsonResult = $stdout | ConvertFrom-Json
                                    $pythonSuccess = $jsonResult.success -eq $true
                                } catch {
                                    $pythonSuccess = $stdout -match '"success".*true'
                                }
                            }

                            # Success if either output file exists with content OR Python reported success
                            $success = $outputExists -or $pythonSuccess
                        } else {
                            # Process timed out, kill it
                            $process.Kill()
                            $result = "Error: Processing timed out after 5 minutes"
                            $success = $false
                        }

                        # Clean up temp files
                        Remove-Item "$env:TEMP\docling_output.txt" -Force -ErrorAction SilentlyContinue
                        Remove-Item "$env:TEMP\docling_error.txt" -Force -ErrorAction SilentlyContinue

                    } finally {
                        Remove-Item $tempPy -Force -ErrorAction SilentlyContinue
                    }
                } else {
                    # Simulation mode
                    Start-Sleep 2
                    "Simulated conversion of: $($item.FileName)`nGenerated at: $(Get-Date)" | Set-Content $outputFile -Encoding UTF8
                    $success = $true
                }

                if ($success) {
                    Update-ItemStatus $item.Id @{
                        Status = 'Completed'
                        OutputFile = $outputFile
                        EndTime = Get-Date
                    }
                    Write-Host "Completed: $($item.FileName)" -ForegroundColor Green
                } else {
                    # Provide more detailed error information
                    $errorMsg = "Conversion failed. "
                    if (-not $outputExists) {
                        $errorMsg += "Output file not created or empty. "
                    }
                    if (-not $pythonSuccess) {
                        $errorMsg += "Python script did not report success. "
                    }
                    if ($stderr) {
                        $errorMsg += "Python stderr: $stderr"
                    }
                    if ($stdout) {
                        $errorMsg += "Python stdout: $stdout"
                    }
                    throw $errorMsg
                }

            } catch {
                # Capture detailed error information
                $errorDetails = @{
                    ExceptionType = $_.Exception.GetType().Name
                    StackTrace = $_.Exception.StackTrace
                    InnerException = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $null }
                    ScriptStackTrace = $_.ScriptStackTrace
                }

                # Try to get stderr if it exists
                $stderr = ""
                try {
                    if (Test-Path "$env:TEMP\docling_error.txt") {
                        $stderr = Get-Content "$env:TEMP\docling_error.txt" -Raw -ErrorAction SilentlyContinue
                    }
                } catch { }

                Update-ItemStatus $item.Id @{
                    Status = 'Error'
                    Error = $_.Exception.Message
                    ErrorDetails = $errorDetails
                    StdErr = $stderr
                    EndTime = Get-Date
                }
                Write-Host "Error processing $($item.FileName): $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        Start-Sleep 2
    }
}

function Start-APIServer {
    [CmdletBinding()]
    param([int]$Port = 8080)

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$Port/")

    try {
        $listener.Start()
        Write-Host "API Server started on port $Port" -ForegroundColor Green

        while ($listener.IsListening) {
            try {
                $context = $listener.GetContext()
                $request = $context.Request
                $response = $context.Response

                # CORS
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")

                if ($request.HttpMethod -eq "OPTIONS") {
                    $response.StatusCode = 200
                    $response.Close()
                    continue
                }

                $responseContent = ""
                $path = $request.Url.LocalPath

            switch -Regex ($path) {
                '^/api/health$' {
                    $responseContent = @{
                        status = 'healthy'
                        timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                        version = $script:DoclingSystem.Version
                        codeVersion = "2025-09-21-updated"
                    } | ConvertTo-Json
                }

                '^/api/status$' {
                    $queue = Get-QueueItems
                    $allStatus = Get-ProcessingStatus
                    $queued = $allStatus.Values | Where-Object { $_.Status -eq 'Queued' }
                    $processing = $allStatus.Values | Where-Object { $_.Status -eq 'Processing' }
                    $completed = $allStatus.Values | Where-Object { $_.Status -eq 'Completed' }
                    $errors = $allStatus.Values | Where-Object { $_.Status -eq 'Error' }

                    $responseContent = @{
                        QueuedCount = @($queued).Count
                        ProcessingCount = @($processing).Count
                        CompletedCount = @($completed).Count
                        ErrorCount = @($errors).Count
                        TotalItems = $allStatus.Count
                    } | ConvertTo-Json
                }

                '^/api/documents$' {
                    $allStatus = Get-ProcessingStatus
                    $documents = @()

                    if ($allStatus -and $allStatus.Count -gt 0) {
                        $allStatus.GetEnumerator() | ForEach-Object {
                            $documents += @{
                                id = $_.Key
                                fileName = if ($_.Value.FileName) { $_.Value.FileName } else { "unknown" }
                                status = $_.Value.Status
                                exportFormat = if ($_.Value.ExportFormat) { $_.Value.ExportFormat } else { "markdown" }
                            }
                        }
                    }

                    $responseContent = if ($documents.Count -eq 0) { "[]" } else {
                        # Ensure array format even for single item
                        if ($documents.Count -eq 1) {
                            "[$($documents | ConvertTo-Json -Depth 10)]"
                        } else {
                            $documents | ConvertTo-Json -Depth 10
                        }
                    }
                }

                '^/api/files$' {
                    # Use ArrayList to ensure array behavior
                    $files = New-Object System.Collections.ArrayList
                    $processedDir = $script:DoclingSystem.OutputDirectory

                    try {
                        if (Test-Path $processedDir) {
                            Get-ChildItem $processedDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                                $docId = $_.Name
                                try {
                                    Get-ChildItem $_.FullName -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml') } | ForEach-Object {
                                        $filePath = $_.FullName
                                        $fileName = $_.Name
                                        $fileSize = [math]::Round($_.Length / 1KB, 2)
                                        $lastModified = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")

                                        $fileObj = @{
                                            id = $docId
                                            fileName = $fileName
                                            filePath = $filePath
                                            fileSize = "$fileSize KB"
                                            lastModified = $lastModified
                                            downloadUrl = "/api/file/" + $docId + "?file=" + $fileName
                                        }
                                        [void]$files.Add($fileObj)
                                    }
                                } catch {
                                    # Skip directories that can't be accessed
                                    Write-Warning "Could not access directory: $($_.FullName)"
                                }
                            }
                        }

                        # Simple response - frontend handles array conversion
                        $responseContent = if ($files.Count -eq 0) {
                            "[]"
                        } else {
                            $files | ConvertTo-Json -Depth 10
                        }
                    } catch {
                        $response.StatusCode = 500
                        $responseContent = @{
                            error = "Failed to load processed files"
                            details = $_.Exception.Message
                        } | ConvertTo-Json
                    }
                }

                '^/api/error/(.+)$' {
                    $id = $Matches[1]
                    $allStatus = Get-ProcessingStatus
                    $status = $allStatus[$id]

                    try {
                        # Debug information
                        $debugInfo = @{
                            requestedId = $id
                            statusFound = $status -ne $null
                            statusData = $status
                            allStatusKeys = $allStatus.Keys -join ', '
                            statusType = if ($status) { $status.GetType().Name } else { 'null' }
                        }

                        if (-not $status) {
                            $response.StatusCode = 404
                            $responseContent = @{
                                error = "Document not found"
                                debug = $debugInfo
                            } | ConvertTo-Json -Depth 10
                        }
                        elseif ($status.Status -eq 'Error') {
                            $errorDetails = @{
                                id = $id
                                fileName = $status.FileName
                                status = $status.Status
                                error = $status.Error
                                errorDetails = $status.ErrorDetails
                                stderr = $status.StdErr
                                queuedTime = $status.QueuedTime
                                startTime = $status.StartTime
                                endTime = $status.EndTime
                                debug = $debugInfo
                            }
                            $responseContent = $errorDetails | ConvertTo-Json -Depth 10
                        }
                        else {
                            $response.StatusCode = 400
                            $responseContent = @{
                                error = "Document is not in error state"
                                currentStatus = $status.Status
                                debug = $debugInfo
                            } | ConvertTo-Json -Depth 10
                        }
                    } catch {
                        $response.StatusCode = 500
                        $responseContent = @{
                            error = "Failed to retrieve error details"
                            details = $_.Exception.Message
                            debug = $debugInfo
                        } | ConvertTo-Json -Depth 10
                    }
                }

                '^/api/file/(.+)$' {
                    $docId = $Matches[1]
                    $processedDir = $script:DoclingSystem.OutputDirectory
                    $docDir = Join-Path $processedDir $docId

                    try {
                        # Check if a specific file is requested via query parameter
                        $requestedFile = $null
                        if ($request.Url.Query) {
                            $queryParams = [System.Web.HttpUtility]::ParseQueryString($request.Url.Query)
                            $requestedFile = $queryParams['file']
                        }

                        if (Test-Path $docDir) {
                            if ($requestedFile) {
                                # Serve specific file
                                $filePath = Join-Path $docDir $requestedFile
                                if (Test-Path $filePath) {
                                    $outputFile = Get-Item $filePath
                                } else {
                                    $outputFile = $null
                                }
                            } else {
                                # Fallback: serve first available file (for Processing Results)
                                $outputFile = Get-ChildItem $docDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml') } | Select-Object -First 1
                            }

                            if ($outputFile -and $outputFile.Extension -in @('.md', '.html', '.json', '.txt', '.xml')) {
                                $bytes = [System.IO.File]::ReadAllBytes($outputFile.FullName)
                                $contentType = switch ($outputFile.Extension) {
                                    '.md' { 'text/markdown; charset=utf-8' }
                                    '.html' { 'text/html; charset=utf-8' }
                                    '.json' { 'application/json; charset=utf-8' }
                                    '.txt' { 'text/plain; charset=utf-8' }
                                    '.xml' { 'application/xml; charset=utf-8' }
                                    default { 'text/plain; charset=utf-8' }
                                }
                                $response.ContentType = $contentType
                                $response.ContentLength64 = $bytes.Length
                                $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                $response.OutputStream.Close()
                                $response.Close()
                                continue
                            }
                        }

                        $response.StatusCode = 404
                        $responseContent = @{ error = "File not found" } | ConvertTo-Json
                    } catch {
                        $response.StatusCode = 500
                        $responseContent = @{
                            error = "Failed to serve file"
                            details = $_.Exception.Message
                        } | ConvertTo-Json
                    }
                }

                '^/api/upload$' {
                    if ($request.HttpMethod -eq 'POST') {
                        try {
                            $reader = New-Object System.IO.StreamReader($request.InputStream)
                            $body = $reader.ReadToEnd()
                            $data = $body | ConvertFrom-Json

                            # Save file
                            $uploadId = [guid]::NewGuid().ToString()
                            $uploadDir = Join-Path $script:DoclingSystem.TempDirectory $uploadId
                            New-Item -ItemType Directory -Force -Path $uploadDir | Out-Null

                            $filePath = Join-Path $uploadDir $data.fileName
                            [System.IO.File]::WriteAllBytes($filePath, [Convert]::FromBase64String($data.dataBase64))

                            # Queue
                            $queueId = Add-DocumentToQueue -Path $filePath

                            $responseContent = @{
                                success = $true
                                documentId = $queueId
                                message = "Document uploaded and queued"
                            } | ConvertTo-Json

                        } catch {
                            $response.StatusCode = 400
                            $responseContent = @{
                                success = $false
                                error = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }
                }

                '^/api/reprocess$' {
                    if ($request.HttpMethod -eq 'POST') {
                        try {
                            $reader = New-Object System.IO.StreamReader($request.InputStream)
                            $body = $reader.ReadToEnd()
                            $data = $body | ConvertFrom-Json

                            $documentId = $data.documentId
                            $newFormat = $data.exportFormat

                            # Get the current document status
                            $allStatus = Get-ProcessingStatus
                            $currentStatus = $allStatus[$documentId]

                            if (-not $currentStatus) {
                                $response.StatusCode = 404
                                $responseContent = @{
                                    success = $false
                                    error = "Document not found"
                                } | ConvertTo-Json
                            } else {
                                # Create new queue item for reprocessing
                                $reprocessItem = @{
                                    Id = $documentId  # Keep same ID to update existing entry
                                    FilePath = $currentStatus.FilePath
                                    FileName = $currentStatus.FileName
                                    ExportFormat = $newFormat
                                    Status = 'Queued'
                                    QueuedTime = Get-Date
                                    IsReprocess = $true
                                }

                                # Add to queue and update status - preserve existing fields
                                Add-QueueItem $reprocessItem
                                Update-ItemStatus $documentId @{
                                    Status = 'Queued'
                                    ExportFormat = $newFormat
                                    QueuedTime = Get-Date
                                    IsReprocess = $true
                                    # Preserve existing fields that might be needed
                                    FilePath = $currentStatus.FilePath
                                    FileName = $currentStatus.FileName
                                    OriginalQueuedTime = $currentStatus.QueuedTime
                                    CompletedTime = $currentStatus.EndTime
                                }

                                $responseContent = @{
                                    success = $true
                                    documentId = $documentId
                                    message = "Document queued for reprocessing with format: $newFormat"
                                } | ConvertTo-Json
                            }
                        } catch {
                            $response.StatusCode = 400
                            $responseContent = @{
                                success = $false
                                error = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }
                }

                '^/api/start-conversion$' {
                    if ($request.HttpMethod -eq 'POST') {
                        try {
                            $reader = New-Object System.IO.StreamReader($request.InputStream)
                            $body = $reader.ReadToEnd()
                            $data = $body | ConvertFrom-Json

                            $documentId = $data.documentId
                            $exportFormat = $data.exportFormat

                            $success = Start-DocumentConversion -DocumentId $documentId -ExportFormat $exportFormat

                            if ($success) {
                                $responseContent = @{
                                    success = $true
                                    documentId = $documentId
                                    message = "Conversion started"
                                } | ConvertTo-Json
                            } else {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error = "Failed to start conversion"
                                } | ConvertTo-Json
                            }
                        } catch {
                            $response.StatusCode = 400
                            $responseContent = @{
                                success = $false
                                error = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }
                }

                '^/api/result/(.+)$' {
                    $id = $Matches[1]
                    $allStatus = Get-ProcessingStatus
                    $status = $allStatus[$id]

                    try {
                        if (-not $status) {
                            $response.StatusCode = 404
                            $responseContent = @{ error = "Document not found" } | ConvertTo-Json
                        }
                        elseif ($status.Status -eq 'Completed' -and $status.OutputFile -and (Test-Path $status.OutputFile)) {
                            $bytes = [System.IO.File]::ReadAllBytes($status.OutputFile)
                            $fileExtension = [System.IO.Path]::GetExtension($status.OutputFile)
                            $contentType = switch ($fileExtension) {
                                '.md' { 'text/markdown; charset=utf-8' }
                                '.html' { 'text/html; charset=utf-8' }
                                '.json' { 'application/json; charset=utf-8' }
                                '.txt' { 'text/plain; charset=utf-8' }
                                '.xml' { 'application/xml; charset=utf-8' }
                                default { 'text/plain; charset=utf-8' }
                            }
                            $response.ContentType = $contentType
                            $response.ContentLength64 = $bytes.Length
                            $response.OutputStream.Write($bytes, 0, $bytes.Length)
                            $response.OutputStream.Close()
                            $response.Close()
                            continue
                        }
                        elseif ($status.Status -in @('Queued','Processing')) {
                            $response.StatusCode = 202
                            $responseContent = @{ status = $status.Status } | ConvertTo-Json
                        }
                        else {
                            $response.StatusCode = 500
                            $responseContent = @{ status = $status.Status; error = $status.Error } | ConvertTo-Json
                        }
                    } catch {
                        $response.StatusCode = 500
                        $responseContent = @{
                            error = "Failed to serve result"
                            details = $_.Exception.Message
                        } | ConvertTo-Json
                    }
                }

                default {
                    $response.StatusCode = 404
                    $responseContent = @{ error = "Not found" } | ConvertTo-Json
                }
            }

                $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                $response.ContentType = "application/json"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.Close()
            } catch {
                # Handle any unhandled exceptions in request processing
                Write-Warning "API Server error processing request: $($_.Exception.Message)"
                try {
                    if ($response -and -not $response.OutputStream.CanWrite) {
                        # Response already closed, skip
                        continue
                    }
                    $errorResponse = @{
                        error = "Internal server error"
                        details = $_.Exception.Message
                    } | ConvertTo-Json
                    $errorBuffer = [System.Text.Encoding]::UTF8.GetBytes($errorResponse)
                    $response.StatusCode = 500
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $errorBuffer.Length
                    $response.OutputStream.Write($errorBuffer, 0, $errorBuffer.Length)
                    $response.Close()
                } catch {
                    # If even error handling fails, just continue
                    Write-Warning "Failed to send error response: $($_.Exception.Message)"
                }
            }
        }
    } finally {
        $listener.Stop()
    }
}

function New-FrontendFiles {
    [CmdletBinding()]
    param()

    $frontendDir = ".\DoclingFrontend"
    if (-not (Test-Path $frontendDir)) {
        New-Item -ItemType Directory -Path $frontendDir -Force | Out-Null
    }

    # Simple HTML file
    $html = @'
<!DOCTYPE html>
<html>
<head>
    <title>Docling Processor</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: linear-gradient(135deg, #1a1a1a 0%, #2d2d2d 100%);
            color: #e0e0e0;
            min-height: 100vh;
        }
        .header {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            padding: 25px;
            border-radius: 12px;
            margin-bottom: 25px;
            border: 1px solid #404040;
            box-shadow: 0 4px 15px rgba(4, 159, 217, 0.1);
        }
        .header h1 {
            margin: 0 0 10px 0;
            background: linear-gradient(45deg, #049fd9, #66d9ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            font-size: 2.2em;
            font-weight: 600;
        }
        .upload-area {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            border: 2px dashed #555;
            border-radius: 12px;
            padding: 40px;
            text-align: center;
            margin-bottom: 25px;
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        .upload-area::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(4, 159, 217, 0.1), transparent);
            transition: left 0.5s;
        }
        .upload-area:hover {
            border-color: #049fd9;
            background: linear-gradient(135deg, #2e2e2e 0%, #3e3e3e 100%);
            box-shadow: 0 4px 20px rgba(4, 159, 217, 0.2);
        }
        .upload-area:hover::before {
            left: 100%;
        }
        .btn {
            background: linear-gradient(135deg, #049fd9 0%, #0284c7 100%);
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 500;
            transition: all 0.3s ease;
            box-shadow: 0 2px 10px rgba(4, 159, 217, 0.3);
        }
        .btn:hover {
            background: linear-gradient(135deg, #0284c7 0%, #049fd9 100%);
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(4, 159, 217, 0.4);
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 25px;
        }
        .stat {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            padding: 25px;
            border-radius: 12px;
            text-align: center;
            border: 1px solid #404040;
            transition: all 0.3s ease;
        }
        .stat:hover {
            transform: translateY(-3px);
            box-shadow: 0 6px 20px rgba(4, 159, 217, 0.15);
            border-color: #049fd9;
        }
        .stat-value {
            font-size: 2.5em;
            font-weight: 700;
            background: linear-gradient(45deg, #049fd9, #66d9ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 5px;
        }
        .stat div:last-child {
            color: #b0b0b0;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .results {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            border-radius: 12px;
            padding: 25px;
            border: 1px solid #404040;
        }
        .results h3 {
            color: #049fd9;
            margin-top: 0;
            font-size: 1.3em;
            font-weight: 600;
        }
        .result-item {
            padding: 15px;
            border-bottom: 1px solid #404040;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-radius: 8px;
            margin-bottom: 8px;
            transition: all 0.3s ease;
        }
        .result-item:hover {
            background: rgba(4, 159, 217, 0.1);
            border-color: #049fd9;
        }
        .result-item:last-child {
            border-bottom: none;
            margin-bottom: 0;
        }
        .result-item strong {
            color: #ffffff;
        }
        .result-item a {
            color: #049fd9;
            text-decoration: none;
            padding: 6px 12px;
            border: 1px solid #049fd9;
            border-radius: 6px;
            transition: all 0.3s ease;
            font-size: 0.9em;
        }
        .result-item a:hover {
            background: #049fd9;
            color: white;
            transform: scale(1.05);
        }
        .format-selector {
            background: #1a1a1a;
            border: 1px solid #555;
            color: #e0e0e0;
            padding: 4px 8px;
            border-radius: 4px;
            margin: 0 8px;
            font-size: 0.85em;
        }
        .format-selector:focus {
            border-color: #049fd9;
            outline: none;
        }
        .reprocess-btn {
            background: #666;
            color: white;
            border: none;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.8em;
            cursor: pointer;
            margin-left: 8px;
            transition: background 0.3s ease;
        }
        .reprocess-btn:hover {
            background: #049fd9;
        }
        .hidden { display: none; }
        .progress {
            width: 100%;
            height: 8px;
            background: #404040;
            border-radius: 10px;
            overflow: hidden;
            margin: 15px 0;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #049fd9, #66d9ff);
            width: 0%;
            transition: width 0.3s ease;
            border-radius: 10px;
        }
        #status {
            color: #049fd9;
            font-weight: 600;
        }
        /* Status indicators */
        .status-ready { color: #049fd9; }
        .status-queued { color: #049fd9; }
        .status-processing { color: #049fd9; }
        .status-completed { color: #10b981; }
        .status-error {
            color: #ef4444;
            cursor: pointer;
            text-decoration: underline;
        }
        .status-error:hover {
            color: #f87171;
        }
        .start-btn {
            background: #3b82f6;
            color: white;
            border: none;
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 0.8em;
            cursor: pointer;
            margin-left: 8px;
            transition: background 0.3s ease;
        }
        .start-btn:hover {
            background: #3b82f6;
        }

        /* Modal styles */
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.7);
        }
        .modal-content {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            margin: 5% auto;
            padding: 30px;
            border: 1px solid #555;
            border-radius: 12px;
            width: 80%;
            max-width: 800px;
            max-height: 80vh;
            overflow-y: auto;
            color: #e0e0e0;
        }
        .modal-header {
            border-bottom: 1px solid #555;
            padding-bottom: 15px;
            margin-bottom: 20px;
        }
        .modal-header h2 {
            margin: 0;
            color: #ef4444;
        }
        .close {
            color: #aaa;
            float: right;
            font-size: 28px;
            font-weight: bold;
            cursor: pointer;
        }
        .close:hover {
            color: #fff;
        }
        .error-section {
            margin-bottom: 20px;
            padding: 15px;
            background: #1a1a1a;
            border-radius: 8px;
            border-left: 4px solid #ef4444;
        }
        .error-section h3 {
            margin-top: 0;
            color: #f87171;
        }
        .error-code {
            background: #0f0f0f;
            padding: 10px;
            border-radius: 6px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            white-space: pre-wrap;
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid #333;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Docling Document Processor</h1>
        <p>Backend Status: <span id="status">Connecting...</span></p>
    </div>

    <div class="upload-area" id="drop-zone">
        <h3>Drop files here or click to browse</h3>
        <p>Supported: PDF, DOCX, XLSX, PPTX, MD, HTML, XHTML, CSV, PNG, JPEG, TIFF, BMP, WEBP</p>
        <input type="file" id="file-input" multiple accept=".pdf,.docx,.xlsx,.pptx,.md,.html,.xhtml,.csv,.png,.jpg,.jpeg,.tiff,.tif,.bmp,.webp" style="display:none">
        <button class="btn" onclick="document.getElementById('file-input').click()">Choose Files</button>
    </div>

    <div id="upload-progress" class="hidden">
        <p>Uploading files...</p>
        <div class="progress"><div class="progress-bar" id="progress-bar"></div></div>
    </div>

    <div class="stats">
        <div class="stat"><div class="stat-value" id="queued">0</div><div>Queued</div></div>
        <div class="stat"><div class="stat-value" id="processing">0</div><div>Processing</div></div>
        <div class="stat"><div class="stat-value" id="completed">0</div><div>Completed</div></div>
        <div class="stat"><div class="stat-value" id="errors">0</div><div>Errors</div></div>
    </div>

    <div class="results">
        <h3>Processing Results</h3>
        <div id="results-list"></div>
    </div>

    <div class="results" style="margin-top: 25px;">
        <h3>Processed Files</h3>
        <div id="files-list">
            <p style="color: #b0b0b0; font-style: italic;">Loading processed files...</p>
        </div>
    </div>

    <!-- Error Details Modal -->
    <div id="errorModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <span class="close">&times;</span>
                <h2>Error Details</h2>
            </div>
            <div id="errorModalContent">
                <p>Loading error details...</p>
            </div>
        </div>
    </div>

    <script>
    const API = 'http://localhost:8080';
    const results = {};

    document.addEventListener('DOMContentLoaded', function() {
        setupUpload();
        // Delay initial API calls to give server time to start
        setTimeout(async () => {
            const isHealthy = await checkHealth();
            if (isHealthy) {
                loadExistingDocuments();
                loadProcessedFiles();
            }
        }, 1000);
        setInterval(loadProcessedFiles, 10000); // Refresh processed files every 10 seconds
        setInterval(updateStats, 2000);
    });

    function setupUpload() {
        const zone = document.getElementById('drop-zone');
        const input = document.getElementById('file-input');

        zone.addEventListener('dragover', e => { e.preventDefault(); zone.style.borderColor = '#007cba'; });
        zone.addEventListener('dragleave', () => { zone.style.borderColor = '#ccc'; });
        zone.addEventListener('drop', e => { e.preventDefault(); zone.style.borderColor = '#ccc'; handleFiles(e.dataTransfer.files); });
        input.addEventListener('change', e => handleFiles(e.target.files));
    }

    async function handleFiles(files) {
        const progress = document.getElementById('upload-progress');
        const bar = document.getElementById('progress-bar');

        progress.classList.remove('hidden');

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                // Use FileReader for reliable base64 conversion
                const base64 = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => {
                        const result = reader.result;
                        const base64Data = result.substring(result.indexOf(',') + 1);
                        resolve(base64Data);
                    };
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                });

                const response = await fetch(API + '/api/upload', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ fileName: file.name, dataBase64: base64 })
                });

                if (response.ok) {
                    const result = await response.json();
                    addResult(result.documentId, file.name);
                    // Don't start polling - document should stay in Ready status for manual conversion
                } else {
                    throw new Error('Upload failed');
                }
            } catch (error) {
                alert('Error uploading ' + file.name + ': ' + error.message);
            }

            bar.style.width = ((i + 1) / files.length * 100) + '%';
        }

        setTimeout(() => { progress.classList.add('hidden'); bar.style.width = '0%'; }, 1000);
    }

    function addResult(id, name, currentFormat = 'markdown') {
        const list = document.getElementById('results-list');
        const item = document.createElement('div');
        item.className = 'result-item';
        item.innerHTML = `
            <div>
                <strong>${name}</strong>
                <span id="status-${id}" class="status-ready">Ready</span>
                <br>
                <small>Format:
                    <select id="format-${id}" class="format-selector" data-id="${id}">
                        <option value="markdown" ${currentFormat === 'markdown' ? 'selected' : ''}>Markdown (.md)</option>
                        <option value="html" ${currentFormat === 'html' ? 'selected' : ''}>HTML (.html)</option>
                        <option value="json" ${currentFormat === 'json' ? 'selected' : ''}>JSON (.json)</option>
                        <option value="text" ${currentFormat === 'text' ? 'selected' : ''}>Plain Text (.txt)</option>
                        <option value="doctags" ${currentFormat === 'doctags' ? 'selected' : ''}>DocTags (.xml)</option>
                    </select>
                    <button class="start-btn" onclick="startConversion('${id}')" id="start-${id}">Start Conversion</button>
                    <button class="reprocess-btn" onclick="reprocessDocument('${id}')" style="display:none" id="reprocess-${id}">Re-process</button>
                </small>
            </div>
            <div>
                <a id="link-${id}" href="#" style="display:none">Download</a>
            </div>
        `;
        list.appendChild(item);
        results[id] = { name: name, format: currentFormat };

        // Show reprocess button when format changes
        document.getElementById(`format-${id}`).addEventListener('change', function() {
            const reprocessBtn = document.getElementById(`reprocess-${id}`);
            const currentStatus = document.getElementById(`status-${id}`).textContent;

            // Only show reprocess button if document is completed
            if (currentStatus === 'Completed') {
                reprocessBtn.style.display = 'inline';
            }
        });
    }

    // Start conversion for a ready document
    async function startConversion(id) {
        const formatSelector = document.getElementById(`format-${id}`);
        const selectedFormat = formatSelector.value;
        const statusElement = document.getElementById(`status-${id}`);
        const startBtn = document.getElementById(`start-${id}`);

        try {
            // Hide start button and update status
            startBtn.style.display = 'none';
            statusElement.textContent = 'Queued...';
            statusElement.className = 'status-queued';

            const response = await fetch(API + '/api/start-conversion', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    documentId: id,
                    exportFormat: selectedFormat
                })
            });

            if (response.ok) {
                const result = await response.json();

                // Update the stored format
                if (results[id] && typeof results[id] === 'object') {
                    results[id].format = selectedFormat;
                } else {
                    results[id] = { name: results[id] || 'Unknown', format: selectedFormat };
                }

                // Start polling for completion
                pollResult(id, results[id].name || 'Unknown');
            } else {
                throw new Error('Failed to start conversion');
            }
        } catch (error) {
            statusElement.textContent = 'Error (click for details)';
            statusElement.className = 'status-error';
            statusElement.onclick = () => showErrorDetails(id, results[id].name || 'Unknown');
            startBtn.style.display = 'inline'; // Show start button again
            console.error('Start conversion error:', error);
        }
    }

    // Re-process document with new format
    async function reprocessDocument(id) {
        const formatSelector = document.getElementById(`format-${id}`);
        const newFormat = formatSelector.value;
        const statusElement = document.getElementById(`status-${id}`);
        const reprocessBtn = document.getElementById(`reprocess-${id}`);

        try {
            // Hide reprocess button and update status
            reprocessBtn.style.display = 'none';
            statusElement.textContent = 'Re-processing...';
            statusElement.className = 'status-processing';

            const response = await fetch(API + '/api/reprocess', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    documentId: id,
                    exportFormat: newFormat
                })
            });

            if (response.ok) {
                const result = await response.json();

                // Update the stored format
                if (results[id] && typeof results[id] === 'object') {
                    results[id].format = newFormat;
                } else {
                    results[id] = { name: results[id] || 'Unknown', format: newFormat };
                }

                // Start polling for completion
                pollResult(id, results[id].name || 'Unknown');
            } else {
                throw new Error('Failed to reprocess document');
            }
        } catch (error) {
            statusElement.textContent = 'Error (click for details)';
            statusElement.className = 'status-error';
            statusElement.onclick = () => showErrorDetails(id, results[id].name || 'Unknown');
            console.error('Reprocess error:', error);
        }
    }

    // Error modal functions
    function showErrorDetails(id, fileName) {
        const modal = document.getElementById('errorModal');
        const content = document.getElementById('errorModalContent');

        content.innerHTML = '<p>Loading error details...</p>';
        modal.style.display = 'block';

        fetch(API + '/api/error/' + id)
            .then(response => {
                console.log('Error API response status:', response.status);
                if (response.ok) {
                    return response.json();
                } else {
                    return response.text().then(text => {
                        console.log('Error API response text:', text);
                        throw new Error(`HTTP ${response.status}: ${text}`);
                    });
                }
            })
            .then(errorData => {
                console.log('Error data received:', errorData);
                content.innerHTML = formatErrorDetails(errorData);
            })
            .catch(error => {
                console.error('Error loading error details:', error);
                content.innerHTML = '<p style="color: #ef4444;">Failed to load error details: ' + error.message + '</p>';
            });
    }

    function formatErrorDetails(errorData) {
        let html = '<div class="error-section">';
        html += '<h3>File Information</h3>';
        html += '<p><strong>File:</strong> ' + (errorData.fileName || 'Unknown') + '</p>';
        html += '<p><strong>Document ID:</strong> ' + errorData.id + '</p>';
        html += '<p><strong>Queued:</strong> ' + (errorData.queuedTime?.DateTime || errorData.queuedTime || 'Unknown') + '</p>';
        html += '<p><strong>Started:</strong> ' + (errorData.startTime?.DateTime || errorData.startTime || 'Unknown') + '</p>';
        html += '<p><strong>Failed:</strong> ' + (errorData.endTime?.DateTime || errorData.endTime || 'Unknown') + '</p>';
        html += '</div>';

        html += '<div class="error-section">';
        html += '<h3>Error Message</h3>';
        html += '<div class="error-code">' + (errorData.error || 'No error message available') + '</div>';
        html += '</div>';

        if (errorData.stderr && errorData.stderr.trim()) {
            html += '<div class="error-section">';
            html += '<h3>Python Error Output (stderr)</h3>';
            html += '<div class="error-code">' + errorData.stderr + '</div>';
            html += '</div>';
        }

        if (errorData.errorDetails) {
            html += '<div class="error-section">';
            html += '<h3>Technical Details</h3>';

            if (errorData.errorDetails.ExceptionType) {
                html += '<p><strong>Exception Type:</strong> ' + errorData.errorDetails.ExceptionType + '</p>';
            }

            if (errorData.errorDetails.InnerException) {
                html += '<p><strong>Inner Exception:</strong> ' + errorData.errorDetails.InnerException + '</p>';
            }

            if (errorData.errorDetails.StackTrace) {
                html += '<h4>Stack Trace:</h4>';
                html += '<div class="error-code">' + errorData.errorDetails.StackTrace + '</div>';
            }

            if (errorData.errorDetails.ScriptStackTrace) {
                html += '<h4>Script Stack Trace:</h4>';
                html += '<div class="error-code">' + errorData.errorDetails.ScriptStackTrace + '</div>';
            }

            html += '</div>';
        }

        return html;
    }

    // Setup modal close functionality
    document.addEventListener('DOMContentLoaded', function() {
        const modal = document.getElementById('errorModal');
        const closeBtn = document.querySelector('.close');

        closeBtn.onclick = function() {
            modal.style.display = 'none';
        }

        window.onclick = function(event) {
            if (event.target === modal) {
                modal.style.display = 'none';
            }
        }
    });

    async function pollResult(id, name, attempt = 0) {
        try {
            const response = await fetch(API + '/api/result/' + id);
            if (response.status === 200) {
                const blob = await response.blob();
                const url = URL.createObjectURL(blob);
                document.getElementById('status-' + id).textContent = 'Completed';
                const link = document.getElementById('link-' + id);
                link.href = url;
                link.style.display = 'inline';
                return;
            }
            if (response.status === 202) {
                document.getElementById('status-' + id).textContent = 'Processing...';
                setTimeout(() => pollResult(id, name, attempt + 1), 2000);
                return;
            }
            const statusElement = document.getElementById('status-' + id);
            statusElement.textContent = 'Error (click for details)';
            statusElement.className = 'status-error';
            statusElement.onclick = () => showErrorDetails(id, name);
        } catch (error) {
            if (attempt < 30) {
                setTimeout(() => pollResult(id, name, attempt + 1), 2000);
            } else {
                const statusElement = document.getElementById('status-' + id);
                statusElement.textContent = 'Error (click for details)';
                statusElement.className = 'status-error';
                statusElement.onclick = () => showErrorDetails(id, name);
            }
        }
    }

    async function updateStats() {
        try {
            const response = await fetch(API + '/api/status');
            if (response.ok) {
                const stats = await response.json();
                document.getElementById('queued').textContent = stats.QueuedCount || 0;
                document.getElementById('processing').textContent = stats.ProcessingCount || 0;
                document.getElementById('completed').textContent = stats.CompletedCount || 0;
                document.getElementById('errors').textContent = stats.ErrorCount || 0;
            } else {
                console.error('Stats update failed with status:', response.status);
                // Try to reconnect if stats fail
                setTimeout(() => checkHealth(1), 1000);
            }
        } catch (error) {
            console.error('Stats update failed:', error);
            // Try to reconnect if stats fail
            setTimeout(() => checkHealth(1), 1000);
        }
    }

    async function loadExistingDocuments() {
        try {
            const response = await fetch(API + '/api/documents');
            if (response.ok) {
                const documents = await response.json();
                const list = document.getElementById('results-list');
                list.innerHTML = ''; // Clear existing items

                documents.forEach(doc => {
                    const docFormat = doc.exportFormat || 'markdown';
                    addResult(doc.id, doc.fileName, docFormat);
                    results[doc.id] = { name: doc.fileName, format: docFormat };

                    // Set appropriate status and start polling if needed
                    const statusElement = document.getElementById('status-' + doc.id);

                    if (doc.status === 'Ready') {
                        statusElement.textContent = 'Ready';
                        statusElement.className = 'status-ready';
                        statusElement.onclick = null; // Clear any existing error click handler
                        // Show start button for ready items
                        const startBtn = document.getElementById(`start-${doc.id}`);
                        if (startBtn) {
                            startBtn.style.display = 'inline';
                        }
                    } else if (doc.status === 'Completed') {
                        statusElement.textContent = 'Completed';
                        statusElement.className = 'status-completed';
                        statusElement.onclick = null; // Clear any existing error click handler
                        const link = document.getElementById('link-' + doc.id);
                        link.href = API + '/api/result/' + doc.id;
                        link.style.display = 'inline';

                        // Hide start button and show reprocess button for completed items
                        const startBtn = document.getElementById(`start-${doc.id}`);
                        if (startBtn) {
                            startBtn.style.display = 'none';
                        }
                        const reprocessBtn = document.getElementById(`reprocess-${doc.id}`);
                        if (reprocessBtn) {
                            reprocessBtn.style.display = 'none'; // Initially hidden until format changes
                        }
                    } else if (doc.status === 'Processing') {
                        statusElement.textContent = 'Processing...';
                        statusElement.className = 'status-processing';
                        statusElement.onclick = null; // Clear any existing error click handler
                        // Hide start button during processing
                        const startBtn = document.getElementById(`start-${doc.id}`);
                        if (startBtn) {
                            startBtn.style.display = 'none';
                        }
                        pollResult(doc.id, doc.fileName);
                    } else if (doc.status === 'Queued') {
                        statusElement.textContent = 'Queued...';
                        statusElement.className = 'status-queued';
                        statusElement.onclick = null; // Clear any existing error click handler
                        // Hide start button when queued
                        const startBtn = document.getElementById(`start-${doc.id}`);
                        if (startBtn) {
                            startBtn.style.display = 'none';
                        }
                        pollResult(doc.id, doc.fileName);
                    } else if (doc.status === 'Error') {
                        statusElement.textContent = 'Error (click for details)';
                        statusElement.className = 'status-error';
                        statusElement.onclick = () => showErrorDetails(doc.id, doc.fileName);
                        // Show start button again for retry
                        const startBtn = document.getElementById(`start-${doc.id}`);
                        if (startBtn) {
                            startBtn.style.display = 'inline';
                        }
                    }
                });
            }
        } catch (error) {
            console.error('Failed to load existing documents:', error);
        }
    }

    async function loadProcessedFiles() {
        try {
            const response = await fetch(API + '/api/files');
            if (response.ok) {
                let files = await response.json();

                // PowerShell API returns single object when only one file exists
                // Convert to array for consistent handling
                if (!Array.isArray(files)) {
                    files = [files];
                }

                const filesList = document.getElementById('files-list');

                if (files.length === 0) {
                    filesList.innerHTML = '<p style="color: #b0b0b0; font-style: italic;">No processed files found.</p>';
                    return;
                }

                filesList.innerHTML = files.map(file => `
                    <div class="result-item">
                        <div>
                            <strong>${file.fileName}</strong><br>
                            <small style="color: #b0b0b0;">
                                Size: ${file.fileSize} | Modified: ${file.lastModified}
                            </small>
                        </div>
                        <div>
                            <a href="${API}${file.downloadUrl}" target="_blank">Download</a>
                        </div>
                    </div>
                `).join('');
            } else {
                console.error('Server responded with status:', response.status);
                const filesList = document.getElementById('files-list');
                filesList.innerHTML = '<p style="color: #fbbf24;">Server responded with an error. Checking connection...</p>';
                // Try to reconnect
                setTimeout(() => checkHealth(1), 2000);
            }
        } catch (error) {
            console.error('Failed to load processed files:', error);
            document.getElementById('files-list').innerHTML =
                '<p style="color: #fbbf24;">Connection lost. Attempting to reconnect...</p>';
            // Try to reconnect
            setTimeout(() => checkHealth(1), 2000);
        }
    }

    async function checkHealth(retries = 3) {
        try {
            const response = await fetch(API + '/api/health', { timeout: 5000 });
            if (response.ok) {
                document.getElementById('status').textContent = 'Connected';
                // If we just reconnected, refresh the processed files
                if (document.getElementById('files-list').innerHTML.includes('Connection lost') ||
                    document.getElementById('files-list').innerHTML.includes('Server responded with an error')) {
                    loadProcessedFiles();
                }
                return true;
            } else {
                document.getElementById('status').textContent = 'Error';
                return false;
            }
        } catch (error) {
            if (retries > 0) {
                document.getElementById('status').textContent = 'Connecting...';
                await new Promise(resolve => setTimeout(resolve, 2000));
                return checkHealth(retries - 1);
            } else {
                document.getElementById('status').textContent = 'Disconnected';
                return false;
            }
        }
    }
    </script>
</body>
</html>
'@

    $html | Set-Content (Join-Path $frontendDir "index.html") -Encoding UTF8

    # Web server script
    $webServer = @'
param([int]$Port = 8081)

$http = New-Object System.Net.HttpListener
$http.Prefixes.Add("http://localhost:$Port/")
$http.Start()

Write-Host "Web server running at http://localhost:$Port" -ForegroundColor Green

try {
    while ($http.IsListening) {
        $context = $http.GetContext()
        $response = $context.Response

        $path = $context.Request.Url.LocalPath
        if ($path -eq "/") { $path = "/index.html" }

        $filePath = Join-Path $PSScriptRoot $path.TrimStart('/')

        if (Test-Path $filePath) {
            $content = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentType = "text/html"
            $response.ContentLength64 = $content.Length
            $response.OutputStream.Write($content, 0, $content.Length)
        } else {
            $response.StatusCode = 404
        }

        $response.Close()
    }
} finally {
    $http.Stop()
}
'@

    $webServer | Set-Content (Join-Path $frontendDir "Start-WebServer.ps1") -Encoding UTF8

    Write-Host "Frontend files created in $frontendDir" -ForegroundColor Green
}

function Start-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$OpenBrowser
    )

    Write-Host "Starting Docling System..." -ForegroundColor Cyan

    # Start API server
    $apiScript = "Import-Module '$PSCommandPath' -Force; Start-APIServer -Port $($script:DoclingSystem.APIPort)"
    $apiPath = Join-Path $env:TEMP "docling_api.ps1"
    $apiScript | Set-Content $apiPath -Encoding UTF8

    $apiProcess = Start-Process powershell -ArgumentList "-File", $apiPath -PassThru -WindowStyle Hidden
    Write-Host "API server started on port $($script:DoclingSystem.APIPort)" -ForegroundColor Green

    # Start processor
    $procScript = "Import-Module '$PSCommandPath' -Force; Start-DocumentProcessor"
    $procPath = Join-Path $env:TEMP "docling_processor.ps1"
    $procScript | Set-Content $procPath -Encoding UTF8

    $procProcess = Start-Process powershell -ArgumentList "-File", $procPath -PassThru -WindowStyle Hidden
    Write-Host "Document processor started" -ForegroundColor Green

    # Start web server
    $webPath = ".\DoclingFrontend\Start-WebServer.ps1"
    if (Test-Path $webPath) {
        $webProcess = Start-Process powershell -ArgumentList "-File", $webPath, "-Port", $script:DoclingSystem.WebPort -PassThru -WindowStyle Hidden
        Write-Host "Web server started on port $($script:DoclingSystem.WebPort)" -ForegroundColor Green

        if ($OpenBrowser) {
            Start-Sleep 2
            Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
        }
    }

    Write-Host "System running! Frontend: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green

    return @{
        API = $apiProcess
        Processor = $procProcess
        Web = $webProcess
    }
}

function Get-DoclingSystemStatus {
    $queue = Get-QueueItems
    $allStatus = Get-ProcessingStatus
    $processing = $allStatus.Values | Where-Object { $_.Status -eq 'Processing' }
    $completed = $allStatus.Values | Where-Object { $_.Status -eq 'Completed' }

    # Test API
    $apiHealthy = $false
    try {
        $response = Invoke-WebRequest -Uri "http://localhost:$($script:DoclingSystem.APIPort)/api/health" -UseBasicParsing -TimeoutSec 2 -ErrorAction SilentlyContinue
        $apiHealthy = $response.StatusCode -eq 200
    } catch {}

    return @{
        Initialized = $true
        Backend = @{
            Running = $true
            ProcessorRunning = $true
            APIHealthy = $apiHealthy
            QueueCount = $queue.Count
            ProcessingCount = @($processing).Count
        }
        Frontend = @{
            Running = $true
            Port = $script:DoclingSystem.WebPort
            URL = "http://localhost:$($script:DoclingSystem.WebPort)"
        }
        System = @{
            Version = $script:DoclingSystem.Version
            TotalDocumentsProcessed = @($completed).Count
        }
    }
}

function Get-PythonStatus {
    return $script:DoclingSystem.PythonAvailable
}

Export-ModuleMember -Function @(
    'Initialize-DoclingSystem',
    'Start-DoclingSystem',
    'Add-DocumentToQueue',
    'Start-DocumentProcessor',
    'Start-APIServer',
    'New-FrontendFiles',
    'Get-DoclingSystemStatus',
    'Get-PythonStatus',
    'Get-QueueItems',
    'Set-QueueItems',
    'Add-QueueItem',
    'Get-NextQueueItem',
    'Get-ProcessingStatus',
    'Set-ProcessingStatus',
    'Update-ItemStatus'
)

Write-Host "Docling System v$($script:DoclingSystem.Version) loaded successfully!" -ForegroundColor Green