<#
.SYNOPSIS
    Use-FileMutex function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Use-FileMutex {
    <#
    .SYNOPSIS
        Executes a script block with exclusive file access using a mutex.

    .DESCRIPTION
        Provides thread-safe and cross-process safe file operations by using a system mutex.
        This prevents race conditions when multiple processes access the same files.

    .PARAMETER Name
        The name of the mutex (should be unique per resource).

    .PARAMETER Script
        The script block to execute with exclusive access.

    .PARAMETER TimeoutMs
        Maximum time to wait for the mutex in milliseconds. Default is 5000 (5 seconds).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$Script,

        [int]$TimeoutMs = 5000
    )

    $created = $false
    $mutex = $null

    try {
        # Create or open a named mutex (Global\ prefix makes it system-wide)
        $mutex = New-Object System.Threading.Mutex($false, "Global\PSDocling_$Name", [ref]$created)

        # Try to acquire the mutex
        if ($mutex.WaitOne($TimeoutMs)) {
            try {
                # Execute the script block with exclusive access
                & $Script
            }
            finally {
                # Always release the mutex
                $mutex.ReleaseMutex() | Out-Null
            }
        }
        else {
            throw "Timeout waiting for mutex: $Name (waited $TimeoutMs ms)"
        }
    }
    finally {
        if ($mutex) {
            $mutex.Dispose()
        }
    }
}
