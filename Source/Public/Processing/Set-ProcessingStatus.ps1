<#
.SYNOPSIS
    Set-ProcessingStatus function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Set-ProcessingStatus {
    param([hashtable]$Status)

    $statusFile = $script:DoclingSystem.StatusFile

    Use-FileMutex -Name "status" -Script {
        # Use atomic write with temp file
        $tempFile = "$statusFile.tmp"
        $Status | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8

        # Atomic move
        Move-Item -Path $tempFile -Destination $statusFile -Force
    }.GetNewClosure()
}
