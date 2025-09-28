<#
.SYNOPSIS
    Get-DoclingSystemStatus function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-DoclingSystemStatus {
    $queue = Get-QueueItems
    $allStatus = Get-ProcessingStatus
    $processing = $allStatus.Values | Where-Object { $_.Status -eq 'Processing' }
    $allCompleted = $allStatus.Values | Where-Object { $_.Status -eq 'Completed' }

    # Calculate session-specific completed count
    # Use the SessionCompletedCount if available (incremented when docs complete)
    # Otherwise calculate based on historical count
    $sessionCompletedCount = 0

    if ($script:DoclingSystem.ContainsKey('SessionCompletedCount')) {
        # Use the session counter that gets incremented in Update-ItemStatus
        $sessionCompletedCount = $script:DoclingSystem.SessionCompletedCount
    } elseif ($script:DoclingSystem.ContainsKey('HistoricalCompletedCount')) {
        # Calculate based on difference from start of session
        $sessionCompletedCount = [Math]::Max(0, @($allCompleted).Count - $script:DoclingSystem.HistoricalCompletedCount)
    } else {
        # If no session tracking, show 0 (fresh session)
        $sessionCompletedCount = 0
    }

    # Test API
    $apiHealthy = $false
    try {
        $response = Invoke-WebRequest -Uri "http://localhost:$($script:DoclingSystem.APIPort)/api/health" -UseBasicParsing -TimeoutSec 2 -ErrorAction SilentlyContinue
        $apiHealthy = $response.StatusCode -eq 200
    }
    catch {}

    return @{
        Initialized = $true
        Backend     = @{
            Running          = $true
            ProcessorRunning = $true
            APIHealthy       = $apiHealthy
            QueueCount       = $queue.Count
            ProcessingCount  = @($processing).Count
        }
        Frontend    = @{
            Running = $true
            Port    = $script:DoclingSystem.WebPort
            URL     = "http://localhost:$($script:DoclingSystem.WebPort)"
        }
        System      = @{
            Version                 = $script:DoclingSystem.Version
            TotalDocumentsProcessed = $sessionCompletedCount
            HistoricalTotal         = @($allCompleted).Count
        }
    }
}
