<#
.SYNOPSIS
    Add-DocumentToQueue function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Add-DocumentToQueue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]$Path,
        [ValidateSet('markdown', 'html', 'json', 'text', 'doctags')]
        [string]$ExportFormat = 'markdown',
        [switch]$EmbedImages,
        [switch]$EnrichCode,
        [switch]$EnrichFormula,
        [switch]$EnrichPictureClasses,
        [switch]$EnrichPictureDescription,

        # Hybrid Chunking Parameters
        [switch]$EnableChunking,
        [ValidateSet('hf','openai')]
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
                    Id                       = [guid]::NewGuid().ToString()
                    FilePath                 = $fileInfo.FullName
                    FileName                 = $fileInfo.Name
                    ExportFormat             = $ExportFormat
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
                    ChunkImagePlaceholder    = $ChunkImagePlaceholder
                    ChunkOverlapTokens       = $ChunkOverlapTokens
                    ChunkPreserveSentences   = $ChunkPreserveSentences.IsPresent
                    ChunkPreserveCode        = $ChunkPreserveCode.IsPresent
                    ChunkModelPreset         = $ChunkModelPreset

                    Status                   = 'Ready'
                    UploadedTime             = Get-Date
                }

                # Don't add to processing queue yet - just store status
                Update-ItemStatus $item.Id $item

                Write-Host "Queued: $($fileInfo.Name) (ID: $($item.Id))" -ForegroundColor Green
                Write-Output $item.Id
            }
        }
    }
}
