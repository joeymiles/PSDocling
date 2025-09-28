<#
.SYNOPSIS
    Get-QueueItems function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-QueueItems {
    $queueFile = $script:DoclingSystem.QueueFile

    # Capture the variable for the closure
    $localQueueFile = $queueFile

    $result = Use-FileMutex -Name "queue" -Script {
        $items = @()
        if (Test-Path $localQueueFile) {
            try {
                $content = Get-Content $localQueueFile -Raw
                if ($content.Trim() -ne "[]") {
                    # Force array conversion in PowerShell 5.1
                    $items = @($content | ConvertFrom-Json)
                }
            }
            catch {
                # Return empty array on error
            }
        }
        return $items
    }.GetNewClosure()

    if ($result) { return $result } else { return @() }
}
