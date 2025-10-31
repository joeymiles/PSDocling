<#
.SYNOPSIS
    Adds a document to the folder-based queue
.DESCRIPTION
    Creates a queue file in the queue folder representing a job to process
.NOTES
    Part of PSDocling Document Processing System
#>
function Add-QueueItemFolder {
    param(
        [Parameter(Mandatory)]
        [string]$DocumentId
    )

    # Ensure queue folder exists
    $queueFolder = "$env:TEMP\DoclingQueue"
    if (-not (Test-Path $queueFolder)) {
        New-Item -Path $queueFolder -ItemType Directory -Force | Out-Null
    }

    # Create a queue file for this document
    # File name format: timestamp_documentId.queue
    $timestamp = [DateTime]::Now.ToString("yyyyMMddHHmmssffff")
    $queueFile = Join-Path $queueFolder "${timestamp}_${DocumentId}.queue"

    # Write the document ID to the file (simple content)
    $DocumentId | Set-Content -Path $queueFile -Encoding UTF8

    Write-Verbose "Added to queue: $DocumentId (File: $queueFile)"
    return $queueFile
}