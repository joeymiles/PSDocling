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
                    $queue = @($content | ConvertFrom-Json)
                }
            }
            catch {
                $queue = @()
            }
        }

        # Add new item
        $newQueue = @($queue) + @($localItem)

        # Write back atomically
        $tempFile = "$localQueueFile.tmp"
        if ($newQueue.Count -eq 1) {
            "[" + ($newQueue[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
        }
        else {
            $newQueue | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        }
        Move-Item -Path $tempFile -Destination $localQueueFile -Force
    }.GetNewClosure()
}
