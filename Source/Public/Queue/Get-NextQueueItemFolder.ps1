<#
.SYNOPSIS
    Gets the next item from the folder-based queue
.DESCRIPTION
    Finds the oldest queue file in the queue folder and returns its document ID.
    Uses a named mutex so only one process can claim a queue file at a time.
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-NextQueueItemFolder {
    $queueFolder = "$env:TEMP\DoclingQueue"
    $localQueueFolder = $queueFolder

    $result = Use-FileMutex -Name "queuefolder" -Script {
        # Ensure queue folder exists
        if (-not (Test-Path $localQueueFolder)) {
            New-Item -Path $localQueueFolder -ItemType Directory -Force | Out-Null
            return $null
        }

        # Get oldest queue file (FIFO by CreationTime)
        $queueFile = Get-ChildItem -Path $localQueueFolder -Filter "*.queue" -ErrorAction SilentlyContinue |
                     Sort-Object CreationTime |
                     Select-Object -First 1

        if (-not $queueFile) {
            Write-Verbose "No items in queue folder"
            return $null
        }

        try {
            $documentId = (Get-Content -Path $queueFile.FullName -Raw -Encoding UTF8).Trim()
            # Claim by deleting under mutex; if delete fails, another claim won
            Remove-Item -Path $queueFile.FullName -Force -ErrorAction Stop
            Write-Verbose "Retrieved from queue: $documentId (File: $($queueFile.Name))"
            return $documentId
        }
        catch {
            Write-Verbose "Failed to claim queue file $($queueFile.FullName): $($_.Exception.Message)"
            return $null
        }
    }.GetNewClosure()

    return $result
}
