<#
.SYNOPSIS
    Gets all items in the folder-based queue
.DESCRIPTION
    Returns an array of document IDs currently in the queue folder
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-QueueItemsFolder {
    $queueFolder = "$env:TEMP\DoclingQueue"

    # Ensure queue folder exists
    if (-not (Test-Path $queueFolder)) {
        return @()
    }

    # Get all queue files
    $queueFiles = Get-ChildItem -Path $queueFolder -Filter "*.queue" |
                  Sort-Object CreationTime

    if (-not $queueFiles) {
        return @()
    }

    # Read document IDs from all queue files
    $queueItems = @()
    foreach ($file in $queueFiles) {
        $documentId = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        $queueItems += $documentId.Trim()
    }

    Write-Verbose "Found $($queueItems.Count) items in queue"
    return $queueItems
}