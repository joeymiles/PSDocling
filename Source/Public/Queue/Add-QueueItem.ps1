<#
.SYNOPSIS
    Add-QueueItem function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Add-QueueItem {
    param($Item)

    $queueFile = $script:DoclingSystem.QueueFile

    # Capture variables for the closure (similar to Get-NextQueueItem pattern)
    $localQueueFile = $queueFile
    $localItem = $Item

    Use-FileMutex -Name "queue" -Script {
        # Read current queue
        $queue = @()
        if (Test-Path $localQueueFile) {
            try {
                $content = Get-Content $localQueueFile -Raw
                if ($content.Trim() -ne "[]") {
                    $parsed = $content | ConvertFrom-Json
                    # Ensure we get an array
                    if ($parsed -is [array]) {
                        $queue = $parsed
                    } else {
                        $queue = @($parsed)
                    }
                }
            }
            catch {
                $queue = @()
            }
        }

        # Add new item - ensure both are arrays before concatenating
        $queue = @($queue)
        $newQueue = $queue + $localItem

        # Write back atomically - ALWAYS force as an array with explicit formatting
        $tempFile = "$localQueueFile.tmp"
        if ($newQueue.Count -eq 0) {
            "[]" | Set-Content $tempFile -Encoding UTF8
        }
        elseif ($newQueue.Count -eq 1) {
            # Force single item to be an array in JSON
            "[" + ($newQueue[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
        }
        else {
            @($newQueue) | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        }
        Move-Item -Path $tempFile -Destination $localQueueFile -Force
    }.GetNewClosure()
}