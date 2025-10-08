<#
.SYNOPSIS
    Start-APIServer function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
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
                            status      = 'healthy'
                            timestamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                            version     = $script:DoclingSystem.Version
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
                            QueuedCount     = @($queued).Count
                            ProcessingCount = @($processing).Count
                            CompletedCount  = @($completed).Count
                            ErrorCount      = @($errors).Count
                            TotalItems      = $allStatus.Count
                        } | ConvertTo-Json
                    }

                    '^/api/documents$' {
                        $allStatus = Get-ProcessingStatus
                        $documents = @()

                        if ($allStatus -and $allStatus.Count -gt 0) {
                            $allStatus.GetEnumerator() | ForEach-Object {
                                $doc = @{
                                    id           = $_.Key
                                    fileName     = if ($_.Value.FileName) { $_.Value.FileName } else { "unknown" }
                                    status       = $_.Value.Status
                                    exportFormat = if ($_.Value.ExportFormat) { $_.Value.ExportFormat } else { "markdown" }
                                }

                                # Add enrichment options - always include with default false if not set
                                $doc.embedImages = if ($_.Value.EmbedImages) { [bool]$_.Value.EmbedImages } else { $false }
                                $doc.enrichCode = if ($_.Value.EnrichCode) { [bool]$_.Value.EnrichCode } else { $false }
                                $doc.enrichFormula = if ($_.Value.EnrichFormula) { [bool]$_.Value.EnrichFormula } else { $false }
                                $doc.enrichPictureClasses = if ($_.Value.EnrichPictureClasses) { [bool]$_.Value.EnrichPictureClasses } else { $false }
                                $doc.enrichPictureDescription = if ($_.Value.EnrichPictureDescription) { [bool]$_.Value.EnrichPictureDescription } else { $false }

                                # Add chunking options with defaults
                                $doc.enableChunking = if ($_.Value.EnableChunking) { [bool]$_.Value.EnableChunking } else { $false }
                                if ($doc.enableChunking) {
                                    $doc.chunkTokenizerBackend = if ($_.Value.ChunkTokenizerBackend) { $_.Value.ChunkTokenizerBackend } else { 'hf' }
                                    $doc.chunkTokenizerModel = if ($_.Value.ChunkTokenizerModel) { $_.Value.ChunkTokenizerModel } else { 'sentence-transformers/all-MiniLM-L6-v2' }
                                    $doc.chunkOpenAIModel = if ($_.Value.ChunkOpenAIModel) { $_.Value.ChunkOpenAIModel } else { 'gpt-4o-mini' }
                                    $doc.chunkMaxTokens = if ($_.Value.ChunkMaxTokens) { $_.Value.ChunkMaxTokens } else { 512 }
                                    $doc.chunkMergePeers = if ($null -ne $_.Value.ChunkMergePeers) { [bool]$_.Value.ChunkMergePeers } else { $true }
                                    $doc.chunkIncludeContext = if ($_.Value.ChunkIncludeContext) { [bool]$_.Value.ChunkIncludeContext } else { $false }
                                    $doc.chunkTableSerialization = if ($_.Value.ChunkTableSerialization) { $_.Value.ChunkTableSerialization } else { 'triplets' }
                                    $doc.chunkPictureStrategy = if ($_.Value.ChunkPictureStrategy) { $_.Value.ChunkPictureStrategy } else { 'default' }
                                    $doc.chunkImagePlaceholder = if ($_.Value.ChunkImagePlaceholder) { $_.Value.ChunkImagePlaceholder } else { '[IMAGE]' }
                                    $doc.chunkOverlapTokens = if ($_.Value.ChunkOverlapTokens) { $_.Value.ChunkOverlapTokens } else { 0 }
                                    $doc.chunkPreserveSentences = if ($_.Value.ChunkPreserveSentences) { [bool]$_.Value.ChunkPreserveSentences } else { $false }
                                    $doc.chunkPreserveCode = if ($_.Value.ChunkPreserveCode) { [bool]$_.Value.ChunkPreserveCode } else { $false }
                                    $doc.chunkModelPreset = if ($_.Value.ChunkModelPreset) { $_.Value.ChunkModelPreset } else { '' }
                                }

                                # Add progress data if available
                                if ($_.Value.Progress) { $doc.progress = $_.Value.Progress }
                                if ($_.Value.EstimatedDuration) { $doc.estimatedDuration = $_.Value.EstimatedDuration }
                                if ($_.Value.ElapsedTime) { $doc.elapsedTime = $_.Value.ElapsedTime }
                                if ($_.Value.FileSize) { $doc.fileSize = $_.Value.FileSize }
                                if ($_.Value.StartTime) { $doc.startTime = $_.Value.StartTime }
                                if ($_.Value.LastUpdate) { $doc.lastUpdate = $_.Value.LastUpdate }

                                $documents += $doc
                            }
                        }

                        $responseContent = if ($documents.Count -eq 0) { "[]" } else {
                            # Ensure array format even for single item
                            if ($documents.Count -eq 1) {
                                "[$($documents | ConvertTo-Json -Depth 10)]"
                            }
                            else {
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
                                        Get-ChildItem $_.FullName -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp') } | ForEach-Object {
                                            $filePath = $_.FullName
                                            $fileName = $_.Name
                                            $fileSize = [math]::Round($_.Length / 1KB, 2)
                                            $lastModified = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")

                                            $fileObj = @{
                                                id           = $docId
                                                fileName     = $fileName
                                                filePath     = $filePath
                                                fileSize     = "$fileSize KB"
                                                lastModified = $lastModified
                                                downloadUrl  = "/api/file/" + $docId + "?file=" + $fileName
                                            }
                                            [void]$files.Add($fileObj)
                                        }
                                    }
                                    catch {
                                        # Skip directories that can't be accessed
                                        Write-Warning "Could not access directory: $($_.FullName)"
                                    }
                                }
                            }

                            # Simple response - frontend handles array conversion
                            $responseContent = if ($files.Count -eq 0) {
                                "[]"
                            }
                            else {
                                $files | ConvertTo-Json -Depth 10
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to load processed files"
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
                                requestedId   = $id
                                statusFound   = $status -ne $null
                                statusData    = $status
                                allStatusKeys = $allStatus.Keys -join ', '
                                statusType    = if ($status) { $status.GetType().Name } else { 'null' }
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
                                    id           = $id
                                    fileName     = $status.FileName
                                    status       = $status.Status
                                    error        = $status.Error
                                    errorDetails = $status.ErrorDetails
                                    stderr       = $status.StdErr
                                    queuedTime   = $status.QueuedTime
                                    startTime    = $status.StartTime
                                    endTime      = $status.EndTime
                                    debug        = $debugInfo
                                }
                                $responseContent = $errorDetails | ConvertTo-Json -Depth 10
                            }
                            else {
                                # For non-error states, provide helpful status information instead of an error
                                $currentInfo = @{
                                    id                = $id
                                    fileName          = $status.FileName
                                    currentStatus     = $status.Status
                                    message           = switch ($status.Status) {
                                        'Processing' { "Document is currently being processed. Progress: $($status.Progress)%" }
                                        'Queued' { "Document is waiting in the processing queue" }
                                        'Ready' { "Document is ready for processing" }
                                        'Completed' { "Document processing completed successfully" }
                                        default { "Document is in $($status.Status) state" }
                                    }
                                    queuedTime        = $status.QueuedTime
                                    startTime         = $status.StartTime
                                    progress          = $status.Progress
                                    estimatedDuration = $status.EstimatedDuration
                                    elapsedTime       = $status.ElapsedTime
                                }

                                # Return 200 OK with status information instead of 400 error
                                $responseContent = $currentInfo | ConvertTo-Json -Depth 10
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to retrieve error details"
                                details = $_.Exception.Message
                                debug   = $debugInfo
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
                                    # Serve specific file with security validation
                                    try {
                                        $secureFileName = Get-SecureFileName -FileName $requestedFile -AllowedExtensions @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp')
                                        $filePath = Join-Path $docDir $secureFileName

                                        # Additional security check: ensure the resolved path is still within the document directory
                                        $resolvedPath = [System.IO.Path]::GetFullPath($filePath)
                                        $resolvedDocDir = [System.IO.Path]::GetFullPath($docDir)

                                        if (-not $resolvedPath.StartsWith($resolvedDocDir + [System.IO.Path]::DirectorySeparatorChar) -and $resolvedPath -ne $resolvedDocDir) {
                                            throw "Path traversal attempt detected"
                                        }

                                        if (Test-Path $filePath) {
                                            $outputFile = Get-Item $filePath
                                        }
                                        else {
                                            $outputFile = $null
                                        }
                                    }
                                    catch {
                                        Write-Warning "Security violation in file request: $($_.Exception.Message)"
                                        $outputFile = $null
                                    }
                                }
                                else {
                                    # Fallback: serve first available file (for Processing Results)
                                    $outputFile = Get-ChildItem $docDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp') } | Select-Object -First 1
                                }

                                if ($outputFile -and $outputFile.Extension -in @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp')) {
                                    $bytes = [System.IO.File]::ReadAllBytes($outputFile.FullName)
                                    $contentType = switch ($outputFile.Extension) {
                                        '.md' { 'text/markdown; charset=utf-8' }
                                        '.html' { 'text/html; charset=utf-8' }
                                        '.json' { 'application/json; charset=utf-8' }
                                        '.txt' { 'text/plain; charset=utf-8' }
                                        '.xml' { 'application/xml; charset=utf-8' }
                                        '.png' { 'image/png' }
                                        '.jpg' { 'image/jpeg' }
                                        '.jpeg' { 'image/jpeg' }
                                        '.gif' { 'image/gif' }
                                        '.bmp' { 'image/bmp' }
                                        '.tiff' { 'image/tiff' }
                                        '.webp' { 'image/webp' }
                                        default { 'application/octet-stream' }
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
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to serve file"
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

                                # Save file with security validation
                                try {
                                    $secureFileName = Get-SecureFileName -FileName $data.fileName
                                }
                                catch {
                                    $response.StatusCode = 400
                                    $responseContent = @{
                                        success = $false
                                        error   = "Invalid filename: $($_.Exception.Message)"
                                    } | ConvertTo-Json
                                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                    $response.OutputStream.Close()
                                    $response.Close()
                                    continue
                                }

                                # Check file size limit (100MB) - validate before decoding to save resources
                                $base64Length = $data.dataBase64.Length
                                $estimatedSizeBytes = ($base64Length * 3) / 4  # Base64 is ~133% of original size
                                $maxSizeBytes = 100 * 1024 * 1024  # 100MB

                                if ($estimatedSizeBytes -gt $maxSizeBytes) {
                                    $sizeMB = [Math]::Round($estimatedSizeBytes / (1024 * 1024), 2)
                                    $response.StatusCode = 413  # Payload Too Large
                                    $responseContent = @{
                                        success = $false
                                        error   = "File size ($sizeMB MB) exceeds the 100MB limit"
                                    } | ConvertTo-Json
                                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                    $response.OutputStream.Close()
                                    $response.Close()
                                    continue
                                }

                                $uploadId = [guid]::NewGuid().ToString()
                                $uploadDir = Join-Path $script:DoclingSystem.TempDirectory $uploadId
                                New-Item -ItemType Directory -Force -Path $uploadDir | Out-Null

                                $filePath = Join-Path $uploadDir $secureFileName
                                $fileBytes = [Convert]::FromBase64String($data.dataBase64)

                                # Double-check actual file size after decoding
                                if ($fileBytes.Length -gt $maxSizeBytes) {
                                    $sizeMB = [Math]::Round($fileBytes.Length / (1024 * 1024), 2)
                                    $response.StatusCode = 413  # Payload Too Large
                                    $responseContent = @{
                                        success = $false
                                        error   = "File size ($sizeMB MB) exceeds the 100MB limit"
                                    } | ConvertTo-Json
                                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                    $response.OutputStream.Close()
                                    $response.Close()
                                    continue
                                }

                                [System.IO.File]::WriteAllBytes($filePath, $fileBytes)

                                # Prepare Add-DocumentToQueue parameters
                                $queueParams = @{
                                    Path = $filePath
                                }

                                # Add optional parameters if provided
                                if ($data.exportFormat) { $queueParams.ExportFormat = $data.exportFormat }
                                if ($data.embedImages -eq $true) { $queueParams.EmbedImages = $true }
                                if ($data.enrichCode -eq $true) { $queueParams.EnrichCode = $true }
                                if ($data.enrichFormula -eq $true) { $queueParams.EnrichFormula = $true }
                                if ($data.enrichPictureClasses -eq $true) { $queueParams.EnrichPictureClasses = $true }
                                if ($data.enrichPictureDescription -eq $true) { $queueParams.EnrichPictureDescription = $true }

                                # Add chunking parameters if provided
                                if ($data.enableChunking -eq $true) { $queueParams.EnableChunking = $true }
                                if ($data.chunkTokenizerBackend) { $queueParams.ChunkTokenizerBackend = $data.chunkTokenizerBackend }
                                if ($data.chunkTokenizerModel) { $queueParams.ChunkTokenizerModel = $data.chunkTokenizerModel }
                                if ($data.chunkOpenAIModel) { $queueParams.ChunkOpenAIModel = $data.chunkOpenAIModel }
                                if ($data.chunkMaxTokens) { $queueParams.ChunkMaxTokens = $data.chunkMaxTokens }
                                if ($null -ne $data.chunkMergePeers) { $queueParams.ChunkMergePeers = $data.chunkMergePeers }
                                if ($data.chunkIncludeContext -eq $true) { $queueParams.ChunkIncludeContext = $true }
                                if ($data.chunkTableSerialization) { $queueParams.ChunkTableSerialization = $data.chunkTableSerialization }
                                if ($data.chunkPictureStrategy) { $queueParams.ChunkPictureStrategy = $data.chunkPictureStrategy }
                                if ($data.chunkImagePlaceholder) { $queueParams.ChunkImagePlaceholder = $data.chunkImagePlaceholder }
                                if ($data.chunkOverlapTokens) { $queueParams.ChunkOverlapTokens = $data.chunkOverlapTokens }
                                if ($data.chunkPreserveSentences -eq $true) { $queueParams.ChunkPreserveSentences = $true }
                                if ($data.chunkPreserveCode -eq $true) { $queueParams.ChunkPreserveCode = $true }
                                if ($data.chunkModelPreset) { $queueParams.ChunkModelPreset = $data.chunkModelPreset }

                                # Queue with parameters - ensure we only capture the ID string
                                $queueId = Add-DocumentToQueue @queueParams | Select-Object -Last 1

                                # Ensure we have a string ID
                                if ($queueId -is [string]) {
                                    $documentId = $queueId
                                } else {
                                    # If it's not a string, try to extract the ID
                                    $documentId = if ($queueId.Id) { $queueId.Id } else { $queueId.ToString() }
                                }

                                $responseContent = @{
                                    success    = $true
                                    documentId = $documentId
                                    message    = "Document uploaded and queued"
                                } | ConvertTo-Json

                            }
                            catch {
                                $response.StatusCode = 400
                                # Add more detailed error information for debugging
                                $errorMessage = if ($_.Exception) {
                                    $_.Exception.Message
                                } else {
                                    $_.ToString()
                                }
                                $responseContent = @{
                                    success = $false
                                    error   = $errorMessage
                                    details = $_.ScriptStackTrace
                                } | ConvertTo-Json -Depth 3
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
                                $embedImages = if ($data.embedImages) { $data.embedImages } else { $false }
                                $enrichCode = if ($data.enrichCode) { $data.enrichCode } else { $false }
                                $enrichFormula = if ($data.enrichFormula) { $data.enrichFormula } else { $false }
                                $enrichPictureClasses = if ($data.enrichPictureClasses) { $data.enrichPictureClasses } else { $false }
                                $enrichPictureDescription = if ($data.enrichPictureDescription) { $data.enrichPictureDescription } else { $false }

                                # Chunking parameters
                                $enableChunking = if ($data.enableChunking) { $data.enableChunking } else { $false }
                                $chunkTokenizerBackend = if ($data.chunkTokenizerBackend) { $data.chunkTokenizerBackend } else { 'hf' }
                                $chunkTokenizerModel = if ($data.chunkTokenizerModel) { $data.chunkTokenizerModel } else { 'sentence-transformers/all-MiniLM-L6-v2' }
                                $chunkOpenAIModel = if ($data.chunkOpenAIModel) { $data.chunkOpenAIModel } else { 'gpt-4o-mini' }
                                $chunkMaxTokens = if ($data.chunkMaxTokens) { $data.chunkMaxTokens } else { 512 }
                                $chunkMergePeers = if ($null -ne $data.chunkMergePeers) { $data.chunkMergePeers } else { $true }
                                $chunkIncludeContext = if ($data.chunkIncludeContext) { $data.chunkIncludeContext } else { $false }
                                $chunkTableSerialization = if ($data.chunkTableSerialization) { $data.chunkTableSerialization } else { 'triplets' }
                                $chunkPictureStrategy = if ($data.chunkPictureStrategy) { $data.chunkPictureStrategy } else { 'default' }
                                $chunkImagePlaceholder = if ($data.chunkImagePlaceholder) { $data.chunkImagePlaceholder } else { '[IMAGE]' }
                                $chunkOverlapTokens = if ($data.chunkOverlapTokens) { $data.chunkOverlapTokens } else { 0 }
                                $chunkPreserveSentences = if ($data.chunkPreserveSentences) { $data.chunkPreserveSentences } else { $false }
                                $chunkPreserveCode = if ($data.chunkPreserveCode) { $data.chunkPreserveCode } else { $false }
                                $chunkModelPreset = if ($data.chunkModelPreset) { $data.chunkModelPreset } else { '' }

                                # Get the current document status
                                $allStatus = Get-ProcessingStatus
                                $currentStatus = $allStatus[$documentId]

                                if (-not $currentStatus) {
                                    $response.StatusCode = 404
                                    $responseContent = @{
                                        success = $false
                                        error   = "Document not found"
                                    } | ConvertTo-Json
                                }
                                else {
                                    # Create new queue item for reprocessing
                                    $reprocessItem = @{
                                        Id                       = $documentId  # Keep same ID to update existing entry
                                        FilePath                 = $currentStatus.FilePath
                                        FileName                 = $currentStatus.FileName
                                        ExportFormat             = $newFormat
                                        EmbedImages              = $embedImages
                                        EnrichCode               = $enrichCode
                                        EnrichFormula            = $enrichFormula
                                        EnrichPictureClasses     = $enrichPictureClasses
                                        EnrichPictureDescription = $enrichPictureDescription

                                        # Chunking Options
                                        EnableChunking           = $enableChunking
                                        ChunkTokenizerBackend    = $chunkTokenizerBackend
                                        ChunkTokenizerModel      = $chunkTokenizerModel
                                        ChunkOpenAIModel         = $chunkOpenAIModel
                                        ChunkMaxTokens           = $chunkMaxTokens
                                        ChunkMergePeers          = $chunkMergePeers
                                        ChunkIncludeContext      = $chunkIncludeContext
                                        ChunkTableSerialization  = $chunkTableSerialization
                                        ChunkPictureStrategy     = $chunkPictureStrategy
                                        ChunkImagePlaceholder    = $chunkImagePlaceholder
                                        ChunkOverlapTokens       = $chunkOverlapTokens
                                        ChunkPreserveSentences   = $chunkPreserveSentences
                                        ChunkPreserveCode        = $chunkPreserveCode
                                        ChunkModelPreset         = $chunkModelPreset

                                        Status                   = 'Queued'
                                        QueuedTime               = Get-Date
                                        IsReprocess              = $true
                                    }

                                    # Add to queue and update status - preserve existing fields
                                    Add-QueueItem $reprocessItem
                                    Update-ItemStatus $documentId @{
                                        Status                   = 'Queued'
                                        ExportFormat             = $newFormat
                                        EmbedImages              = $embedImages
                                        EnrichCode               = $enrichCode
                                        EnrichFormula            = $enrichFormula
                                        EnrichPictureClasses     = $enrichPictureClasses
                                        EnrichPictureDescription = $enrichPictureDescription

                                        # Chunking Options
                                        EnableChunking           = $enableChunking
                                        ChunkTokenizerBackend    = $chunkTokenizerBackend
                                        ChunkTokenizerModel      = $chunkTokenizerModel
                                        ChunkOpenAIModel         = $chunkOpenAIModel
                                        ChunkMaxTokens           = $chunkMaxTokens
                                        ChunkMergePeers          = $chunkMergePeers
                                        ChunkIncludeContext      = $chunkIncludeContext
                                        ChunkTableSerialization  = $chunkTableSerialization
                                        ChunkPictureStrategy     = $chunkPictureStrategy
                                        ChunkImagePlaceholder    = $chunkImagePlaceholder
                                        ChunkOverlapTokens       = $chunkOverlapTokens
                                        ChunkPreserveSentences   = $chunkPreserveSentences
                                        ChunkPreserveCode        = $chunkPreserveCode
                                        ChunkModelPreset         = $chunkModelPreset

                                        QueuedTime               = Get-Date
                                        IsReprocess              = $true
                                        # Preserve existing fields that might be needed
                                        FilePath                 = $currentStatus.FilePath
                                        FileName                 = $currentStatus.FileName
                                        OriginalQueuedTime       = $currentStatus.QueuedTime
                                        CompletedTime            = $currentStatus.EndTime
                                    }

                                    $responseContent = @{
                                        success    = $true
                                        documentId = $documentId
                                        message    = "Document queued for reprocessing with format: $newFormat"
                                    } | ConvertTo-Json
                                }
                            }
                            catch {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error   = $_.Exception.Message
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
                                $embedImages = if ($data.embedImages) { $data.embedImages } else { $false }
                                $enrichCode = if ($data.enrichCode) { $data.enrichCode } else { $false }
                                $enrichFormula = if ($data.enrichFormula) { $data.enrichFormula } else { $false }
                                $enrichPictureClasses = if ($data.enrichPictureClasses) { $data.enrichPictureClasses } else { $false }
                                $enrichPictureDescription = if ($data.enrichPictureDescription) { $data.enrichPictureDescription } else { $false }

                                # Chunking parameters
                                $enableChunking = if ($data.enableChunking) { $data.enableChunking } else { $false }
                                $chunkTokenizerBackend = if ($data.chunkTokenizerBackend) { $data.chunkTokenizerBackend } else { 'hf' }
                                $chunkTokenizerModel = if ($data.chunkTokenizerModel) { $data.chunkTokenizerModel } else { 'sentence-transformers/all-MiniLM-L6-v2' }
                                $chunkOpenAIModel = if ($data.chunkOpenAIModel) { $data.chunkOpenAIModel } else { 'gpt-4o-mini' }
                                $chunkMaxTokens = if ($data.chunkMaxTokens) { $data.chunkMaxTokens } else { 512 }
                                $chunkMergePeers = if ($null -ne $data.chunkMergePeers) { $data.chunkMergePeers } else { $true }
                                $chunkIncludeContext = if ($data.chunkIncludeContext) { $data.chunkIncludeContext } else { $false }
                                $chunkTableSerialization = if ($data.chunkTableSerialization) { $data.chunkTableSerialization } else { 'triplets' }
                                $chunkPictureStrategy = if ($data.chunkPictureStrategy) { $data.chunkPictureStrategy } else { 'default' }
                                $chunkImagePlaceholder = if ($data.chunkImagePlaceholder) { $data.chunkImagePlaceholder } else { '[IMAGE]' }
                                $chunkOverlapTokens = if ($data.chunkOverlapTokens) { $data.chunkOverlapTokens } else { 0 }
                                $chunkPreserveSentences = if ($data.chunkPreserveSentences) { $data.chunkPreserveSentences } else { $false }
                                $chunkPreserveCode = if ($data.chunkPreserveCode) { $data.chunkPreserveCode } else { $false }
                                $chunkModelPreset = if ($data.chunkModelPreset) { $data.chunkModelPreset } else { '' }

                                $success = Start-DocumentConversion -DocumentId $documentId -ExportFormat $exportFormat -EmbedImages:$embedImages -EnrichCode:$enrichCode -EnrichFormula:$enrichFormula -EnrichPictureClasses:$enrichPictureClasses -EnrichPictureDescription:$enrichPictureDescription -EnableChunking:$enableChunking -ChunkTokenizerBackend $chunkTokenizerBackend -ChunkTokenizerModel $chunkTokenizerModel -ChunkOpenAIModel $chunkOpenAIModel -ChunkMaxTokens $chunkMaxTokens -ChunkMergePeers:$chunkMergePeers -ChunkIncludeContext:$chunkIncludeContext -ChunkTableSerialization $chunkTableSerialization -ChunkPictureStrategy $chunkPictureStrategy -ChunkImagePlaceholder $chunkImagePlaceholder -ChunkOverlapTokens $chunkOverlapTokens -ChunkPreserveSentences:$chunkPreserveSentences -ChunkPreserveCode:$chunkPreserveCode -ChunkModelPreset $chunkModelPreset

                                if ($success) {
                                    $responseContent = @{
                                        success    = $true
                                        documentId = $documentId
                                        message    = "Conversion started"
                                    } | ConvertTo-Json
                                }
                                else {
                                    $response.StatusCode = 400
                                    $responseContent = @{
                                        success = $false
                                        error   = "Failed to start conversion"
                                    } | ConvertTo-Json
                                }
                            }
                            catch {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error   = $_.Exception.Message
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
                                $fileInfo = Get-Item $status.OutputFile
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

                                # For large files (>5MB), use streaming
                                if ($fileInfo.Length -gt 5MB) {
                                    Write-Host "Streaming large file: $($fileInfo.Length) bytes"
                                    $response.SendChunked = $true

                                    try {
                                        $fileStream = [System.IO.File]::OpenRead($status.OutputFile)
                                        $buffer = New-Object byte[] 8192

                                        while (($bytesRead = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                                            $response.OutputStream.Write($buffer, 0, $bytesRead)
                                        }
                                    }
                                    finally {
                                        if ($fileStream) { $fileStream.Close() }
                                    }
                                }
                                else {
                                    # For smaller files, use the original method
                                    $bytes = [System.IO.File]::ReadAllBytes($status.OutputFile)
                                    $response.ContentLength64 = $bytes.Length
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                }

                                $response.OutputStream.Close()
                                $response.Close()
                                continue
                            }
                            elseif ($status.Status -in @('Queued', 'Processing')) {
                                $response.StatusCode = 202
                                $responseContent = @{ status = $status.Status } | ConvertTo-Json
                            }
                            else {
                                $response.StatusCode = 500
                                $responseContent = @{ status = $status.Status; error = $status.Error } | ConvertTo-Json
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to serve result"
                                details = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }

                    '^/api/cancel/(.+)$' {
                        $id = $Matches[1]
                        try {
                            # Mark document for cancellation
                            $allStatus = Get-ProcessingStatus
                            $status = $allStatus[$id]

                            if (-not $status) {
                                $response.StatusCode = 404
                                $responseContent = @{
                                    success = $false
                                    error   = "Document not found"
                                } | ConvertTo-Json
                            }
                            elseif ($status.Status -eq 'Processing') {
                                # Update status to mark for cancellation
                                Update-ItemStatus $id @{
                                    CancelRequested = $true
                                }
                                $response.StatusCode = 200
                                $responseContent = @{
                                    success = $true
                                    message = "Cancellation requested"
                                } | ConvertTo-Json
                            }
                            else {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error   = "Document is not currently processing (Status: $($status.Status))"
                                } | ConvertTo-Json
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                success = $false
                                error   = "Failed to cancel document"
                                details = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }

                    '^/api/documents/(.+)/reset$' {
                        if ($request.HttpMethod -eq 'POST') {
                            $documentId = $Matches[1]
                            try {
                                # Get current document status
                                $allStatus = Get-ProcessingStatus
                                $status = $allStatus[$documentId]

                                if (-not $status) {
                                    $response.StatusCode = 404
                                    $responseContent = @{
                                        error      = "Document not found"
                                        documentId = $documentId
                                    } | ConvertTo-Json
                                }
                                else {
                                    # Reset document to Ready status
                                    Update-ItemStatus $documentId @{
                                        Status       = 'Ready'
                                        Progress     = 0
                                        StartTime    = $null
                                        EndTime      = $null
                                        Error        = $null
                                        ErrorDetails = $null
                                        StdErr       = $null
                                        LastUpdate   = Get-Date
                                    }

                                    $responseContent = @{
                                        success    = $true
                                        documentId = $documentId
                                        status     = 'Ready'
                                        message    = "Document status reset to Ready"
                                    } | ConvertTo-Json
                                }
                            }
                            catch {
                                $response.StatusCode = 500
                                $responseContent = @{
                                    error   = "Failed to reset document status"
                                    details = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                    }

                    # Download single document as ZIP
                    '^/api/download/([a-fA-F0-9\-]+)$' {
                        $docId = $Matches[1]
                        $processedDir = $script:DoclingSystem.OutputDirectory
                        $docDir = Join-Path $processedDir $docId

                        if (Test-Path $docDir) {
                            try {
                                # Create ZIP file in temp directory
                                $zipPath = Join-Path $env:TEMP "$docId.zip"

                                # Remove existing ZIP if it exists
                                if (Test-Path $zipPath) {
                                    Remove-Item $zipPath -Force
                                }

                                # Create ZIP archive
                                Compress-Archive -Path "$docDir\*" -DestinationPath $zipPath -Force

                                # Send ZIP file
                                $fileBytes = [System.IO.File]::ReadAllBytes($zipPath)
                                $response.ContentType = "application/zip"
                                $response.ContentLength64 = $fileBytes.Length
                                $response.Headers.Add("Content-Disposition", "attachment; filename=`"$docId.zip`"")
                                $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
                                $response.Close()

                                # Clean up temp file
                                Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
                                continue
                            }
                            catch {
                                $response.StatusCode = 500
                                $responseContent = @{
                                    error = "Failed to create download"
                                    details = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                        else {
                            $response.StatusCode = 404
                            $responseContent = @{ error = "Document not found" } | ConvertTo-Json
                        }
                    }

                    # Download all documents as ZIP
                    '^/api/download-all$' {
                        $processedDir = $script:DoclingSystem.OutputDirectory

                        if (Test-Path $processedDir) {
                            try {
                                # Create timestamp for filename
                                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                                $zipPath = Join-Path $env:TEMP "PSDocling_Export_$timestamp.zip"

                                # Remove existing ZIP if it exists
                                if (Test-Path $zipPath) {
                                    Remove-Item $zipPath -Force
                                }

                                # Get all document folders
                                $docFolders = Get-ChildItem -Path $processedDir -Directory

                                if ($docFolders.Count -eq 0) {
                                    $response.StatusCode = 404
                                    $responseContent = @{ error = "No documents to download" } | ConvertTo-Json
                                }
                                else {
                                    # Create ZIP archive with all documents
                                    Compress-Archive -Path "$processedDir\*" -DestinationPath $zipPath -Force

                                    # Send ZIP file
                                    $fileBytes = [System.IO.File]::ReadAllBytes($zipPath)
                                    $response.ContentType = "application/zip"
                                    $response.ContentLength64 = $fileBytes.Length
                                    $response.Headers.Add("Content-Disposition", "attachment; filename=`"PSDocling_Export_$timestamp.zip`"")
                                    $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
                                    $response.Close()

                                    # Clean up temp file
                                    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
                                    continue
                                }
                            }
                            catch {
                                $response.StatusCode = 500
                                $responseContent = @{
                                    error = "Failed to create bulk download"
                                    details = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                        else {
                            $response.StatusCode = 404
                            $responseContent = @{ error = "No processed documents directory found" } | ConvertTo-Json
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
            }
            catch {
                # Handle any unhandled exceptions in request processing
                Write-Warning "API Server error processing request: $($_.Exception.Message)"
                try {
                    if ($response -and -not $response.OutputStream.CanWrite) {
                        # Response already closed, skip
                        continue
                    }
                    $errorResponse = @{
                        error   = "Internal server error"
                        details = $_.Exception.Message
                    } | ConvertTo-Json
                    $errorBuffer = [System.Text.Encoding]::UTF8.GetBytes($errorResponse)
                    $response.StatusCode = 500
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $errorBuffer.Length
                    $response.OutputStream.Write($errorBuffer, 0, $errorBuffer.Length)
                    $response.Close()
                }
                catch {
                    # If even error handling fails, just continue
                    Write-Warning "Failed to send error response: $($_.Exception.Message)"
                }
            }
        }
    }
    finally {
        $listener.Stop()
    }
}
