<#
.SYNOPSIS
    Adds a document to the folder-based queue
.DESCRIPTION
    Creates a queue file in the queue folder representing a job to process.
    Uses the same mutex as Get-NextQueueItemFolder to avoid races with claimers.
.NOTES
    Part of PSDocling Document Processing System
#>
function Add-QueueItemFolder {
    param(
        [Parameter(Mandatory)]
        [string]$DocumentId
    )

    $queueFolder = "$env:TEMP\DoclingQueue"
    $localQueueFolder = $queueFolder
    $localDocumentId = $DocumentId

    $result = Use-FileMutex -Name "queuefolder" -Script {
        if (-not (Test-Path $localQueueFolder)) {
            New-Item -Path $localQueueFolder -ItemType Directory -Force | Out-Null
        }

        # File name format: timestamp_documentId.queue
        $timestamp = [DateTime]::Now.ToString("yyyyMMddHHmmssffff")
        $queueFile = Join-Path $localQueueFolder "${timestamp}_${localDocumentId}.queue"

        $localDocumentId | Set-Content -Path $queueFile -Encoding UTF8
        Write-Verbose "Added to queue: $localDocumentId (File: $queueFile)"
        return $queueFile
    }.GetNewClosure()

    return $result
}
