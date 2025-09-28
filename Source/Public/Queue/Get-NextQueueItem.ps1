<#
.SYNOPSIS
    Get-NextQueueItem function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-NextQueueItem {
    $queueFile = $script:DoclingSystem.QueueFile

    # Capture variables for the closure
    $localQueueFile = $queueFile

    $result = Use-FileMutex -Name "queue" -Script {
        $nextItem = $null
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

        if ($queue.Count -gt 0) {
            $nextItem = $queue[0]
            $remaining = if ($queue.Count -gt 1) { $queue[1..($queue.Count - 1)] } else { @() }

            # Write remaining items back atomically
            $tempFile = "$localQueueFile.tmp"
            if ($remaining.Count -eq 0) {
                "[]" | Set-Content $tempFile -Encoding UTF8
            }
            elseif ($remaining.Count -eq 1) {
                "[" + ($remaining[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
            }
            else {
                $remaining | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
            }
            Move-Item -Path $tempFile -Destination $localQueueFile -Force
        }

        return $nextItem
    }.GetNewClosure()

    return $result
}
