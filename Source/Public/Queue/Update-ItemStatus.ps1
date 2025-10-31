<#
.SYNOPSIS
    Update-ItemStatus function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Update-ItemStatus {
    param($Id, $Updates)

    $statusFile = $script:DoclingSystem.StatusFile

    # Capture variables for the closure
    $localStatusFile = $statusFile
    $localId = $Id
    $localUpdates = $Updates

    Use-FileMutex -Name "status" -Script {
        # Read current status
        $status = @{}
        if (Test-Path $localStatusFile) {
            try {
                $content = Get-Content $localStatusFile -Raw
                $jsonObj = $content | ConvertFrom-Json

                # Convert PSCustomObject to hashtable manually
                $hashtable = @{}
                $jsonObj.PSObject.Properties | ForEach-Object {
                    $hashtable[$_.Name] = $_.Value
                }
                $status = $hashtable
            }
            catch {
                $status = @{}
            }
        }

        # Convert existing item to hashtable if it's a PSObject
        if ($status[$localId]) {
            if ($status[$localId] -is [PSCustomObject]) {
                $itemHash = @{}
                $status[$localId].PSObject.Properties | ForEach-Object {
                    $itemHash[$_.Name] = $_.Value
                }
                $status[$localId] = $itemHash
            }
            # ENSURE it's a hashtable - if not, create new one with existing properties
            if ($status[$localId] -isnot [hashtable]) {
                $oldItem = $status[$localId]
                $status[$localId] = @{}
                # Try to copy any existing properties
                if ($oldItem -is [PSCustomObject]) {
                    $oldItem.PSObject.Properties | ForEach-Object {
                        $status[$localId][$_.Name] = $_.Value
                    }
                }
            }
        }
        else {
            $status[$localId] = @{}
        }

        # Track session completion count (before applying updates)
        if ($localUpdates.ContainsKey('Status') -and $localUpdates['Status'] -eq 'Completed') {
            # Check if this item wasn't already completed
            $wasCompleted = $status[$localId] -and $status[$localId]['Status'] -eq 'Completed'
            if (-not $wasCompleted) {
                if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('SessionCompletedCount')) {
                    $script:DoclingSystem.SessionCompletedCount++
                }
            }
        }

        # Apply updates
        foreach ($key in $localUpdates.Keys) {
            $status[$localId][$key] = $localUpdates[$key]
        }

        # Write back atomically
        $tempFile = "$localStatusFile.tmp"
        $status | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        Move-Item -Path $tempFile -Destination $localStatusFile -Force

        # Also update local cache (ensure it's initialized)
        if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('ProcessingStatus')) {
            if ($null -eq $script:DoclingSystem.ProcessingStatus) {
                $script:DoclingSystem['ProcessingStatus'] = @{}
            }
            $script:DoclingSystem['ProcessingStatus'][$localId] = $status[$localId]
        }
    }.GetNewClosure()
}
