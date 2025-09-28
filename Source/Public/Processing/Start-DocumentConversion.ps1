<#
.SYNOPSIS
    Start-DocumentConversion function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Start-DocumentConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DocumentId,
        [ValidateSet('markdown', 'html', 'json', 'text', 'doctags')]
        [string]$ExportFormat,
        [switch]$EmbedImages,
        [switch]$EnrichCode,
        [switch]$EnrichFormula,
        [switch]$EnrichPictureClasses,
        [switch]$EnrichPictureDescription,

        # Hybrid Chunking Parameters
        [switch]$EnableChunking,
        [ValidateSet('hf', 'openai')]
        [string]$ChunkTokenizerBackend = 'hf',
        [string]$ChunkTokenizerModel = 'sentence-transformers/all-MiniLM-L6-v2',
        [string]$ChunkOpenAIModel = 'gpt-4o-mini',
        [ValidateRange(50, 8192)]
        [int]$ChunkMaxTokens = 512,
        [bool]$ChunkMergePeers = $true,
        [switch]$ChunkIncludeContext,
        [ValidateSet('triplets', 'markdown', 'csv', 'grid')]
        [string]$ChunkTableSerialization = 'triplets',
        [ValidateSet('default', 'with_caption', 'with_description', 'placeholder')]
        [string]$ChunkPictureStrategy = 'default',
        [string]$ChunkImagePlaceholder = '[IMAGE]',
        [ValidateRange(0, 1000)]
        [int]$ChunkOverlapTokens = 0,
        [switch]$ChunkPreserveSentences,
        [switch]$ChunkPreserveCode,
        [ValidateSet('', 'general', 'legal', 'medical', 'financial', 'scientific', 'multilingual', 'code')]
        [string]$ChunkModelPreset = ''
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
        Id                       = $DocumentId
        FilePath                 = $documentStatus.FilePath
        FileName                 = $documentStatus.FileName
        ExportFormat             = $documentStatus.ExportFormat
        EmbedImages              = $EmbedImages.IsPresent
        EnrichCode               = $EnrichCode.IsPresent
        EnrichFormula            = $EnrichFormula.IsPresent
        EnrichPictureClasses     = $EnrichPictureClasses.IsPresent
        EnrichPictureDescription = $EnrichPictureDescription.IsPresent

        # Chunking Options
        EnableChunking           = $EnableChunking.IsPresent
        ChunkTokenizerBackend    = $ChunkTokenizerBackend
        ChunkTokenizerModel      = $ChunkTokenizerModel
        ChunkOpenAIModel         = $ChunkOpenAIModel
        ChunkMaxTokens           = $ChunkMaxTokens
        ChunkMergePeers          = $ChunkMergePeers
        ChunkIncludeContext      = $ChunkIncludeContext.IsPresent
        ChunkTableSerialization  = $ChunkTableSerialization
        ChunkPictureStrategy     = $ChunkPictureStrategy

        Status                   = 'Queued'
        QueuedTime               = Get-Date
        UploadedTime             = $documentStatus.UploadedTime
    }

    # Add to processing queue and update status
    Add-QueueItem $queueItem
    Update-ItemStatus $DocumentId @{
        Status                   = 'Queued'
        QueuedTime               = Get-Date
        ExportFormat             = $documentStatus.ExportFormat
        EmbedImages              = $EmbedImages.IsPresent
        EnrichCode               = $EnrichCode.IsPresent
        EnrichFormula            = $EnrichFormula.IsPresent
        EnrichPictureClasses     = $EnrichPictureClasses.IsPresent
        EnrichPictureDescription = $EnrichPictureDescription.IsPresent

        # Chunking Options
        EnableChunking           = $EnableChunking.IsPresent
        ChunkTokenizerBackend    = $ChunkTokenizerBackend
        ChunkTokenizerModel      = $ChunkTokenizerModel
        ChunkOpenAIModel         = $ChunkOpenAIModel
        ChunkMaxTokens           = $ChunkMaxTokens
        ChunkMergePeers          = $ChunkMergePeers
        ChunkIncludeContext      = $ChunkIncludeContext.IsPresent
        ChunkTableSerialization  = $ChunkTableSerialization
        ChunkPictureStrategy     = $ChunkPictureStrategy
    }

    Write-Host "Started conversion for: $($documentStatus.FileName) (ID: $DocumentId)" -ForegroundColor Green
    return $true
}
