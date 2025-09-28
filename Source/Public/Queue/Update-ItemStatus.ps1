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

    Use-FileMutex -Name "status" -Script {
        # Read current status
        $status = @{}
        if (Test-Path $statusFile) {
            try {
                $content = Get-Content $statusFile -Raw
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
        if ($status[$Id]) {
            if ($status[$Id] -is [PSCustomObject]) {
                $itemHash = @{}
                $status[$Id].PSObject.Properties | ForEach-Object {
                    $itemHash[$_.Name] = $_.Value
                }
                $status[$Id] = $itemHash
            }
        }
        else {
            $status[$Id] = @{}
        }

        # Track session completion count (before applying updates)
        if ($Updates.ContainsKey('Status') -and $Updates['Status'] -eq 'Completed') {
            # Check if this item wasn't already completed
            $wasCompleted = $status[$Id] -and $status[$Id]['Status'] -eq 'Completed'
            if (-not $wasCompleted) {
                if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('SessionCompletedCount')) {
                    $script:DoclingSystem.SessionCompletedCount++
                }
            }
        }

        # Apply updates
        foreach ($key in $Updates.Keys) {
            $status[$Id][$key] = $Updates[$key]
        }

        # Write back atomically
        $tempFile = "$statusFile.tmp"
        $status | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        Move-Item -Path $tempFile -Destination $statusFile -Force

        # Also update local cache (ensure it's initialized)
        if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('ProcessingStatus')) {
            if ($null -eq $script:DoclingSystem.ProcessingStatus) {
                $script:DoclingSystem['ProcessingStatus'] = @{}
            }
            $script:DoclingSystem['ProcessingStatus'][$Id] = $status[$Id]
        }
    }.GetNewClosure()
}
