<#
.SYNOPSIS
    Get-ProcessingStatus function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-ProcessingStatus {
    $statusFile = $script:DoclingSystem.StatusFile

    # Capture the variable for the closure
    $localStatusFile = $statusFile

    $result = Use-FileMutex -Name "status" -Script {
        $resultHash = @{}
        if (Test-Path $localStatusFile) {
            try {
                $content = Get-Content $localStatusFile -Raw
                $jsonObj = $content | ConvertFrom-Json

                # Convert PSCustomObject to hashtable manually
                $jsonObj.PSObject.Properties | ForEach-Object {
                    $resultHash[$_.Name] = $_.Value
                }
            }
            catch {
                # Return empty hashtable on error
            }
        }
        return $resultHash
    }.GetNewClosure()

    if ($result) { return $result } else { return @{} }
}
