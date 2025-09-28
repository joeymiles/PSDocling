<#
.SYNOPSIS
    Set-QueueItems function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Set-QueueItems {
    param([array]$Items = @())

    $queueFile = $script:DoclingSystem.QueueFile

    # Use local variables that will be captured correctly
    $itemsToWrite = $Items
    $queueFilePath = $queueFile

    Use-FileMutex -Name "queue" -Script {
        # Use atomic write with temp file
        $tempFile = "$queueFilePath.tmp"

        # Ensure we always store as a JSON array, even for single items
        if ($itemsToWrite.Count -eq 0) {
            "[]" | Set-Content $tempFile -Encoding UTF8
        }
        elseif ($itemsToWrite.Count -eq 1) {
            "[" + ($itemsToWrite[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
        }
        else {
            $itemsToWrite | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        }

        # Atomic move
        Move-Item -Path $tempFile -Destination $queueFilePath -Force
    }.GetNewClosure()
}
