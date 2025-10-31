<#
.SYNOPSIS
    Gets the next item from the folder-based queue
.DESCRIPTION
    Finds the oldest queue file in the queue folder and returns its document ID
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-NextQueueItemFolder {
    $queueFolder = "$env:TEMP\DoclingQueue"

    # Ensure queue folder exists
    if (-not (Test-Path $queueFolder)) {
        New-Item -Path $queueFolder -ItemType Directory -Force | Out-Null
        return $null
    }

    # Get all queue files, sorted by creation time (oldest first)
    $queueFiles = Get-ChildItem -Path $queueFolder -Filter "*.queue" |
                  Sort-Object CreationTime |
                  Select-Object -First 1

    if (-not $queueFiles) {
        Write-Verbose "No items in queue folder"
        return $null
    }

    $queueFile = $queueFiles[0]

    # Read the document ID from the file
    $documentId = Get-Content -Path $queueFile.FullName -Raw -Encoding UTF8
    $documentId = $documentId.Trim()

    # Delete the queue file (item is now being processed)
    Remove-Item -Path $queueFile.FullName -Force

    Write-Verbose "Retrieved from queue: $documentId (File: $($queueFile.Name))"
    return $documentId
}