# Manually run processor logic once
Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "=== Manual Processor Test ===" -ForegroundColor Cyan

Write-Host "`n1. Checking queue..." -ForegroundColor Yellow
$queueItems = Get-QueueItems
Write-Host "Queue has $($queueItems.Count) items" -ForegroundColor $(if ($queueItems.Count -gt 0) { 'Green' } else { 'Red' })

if ($queueItems.Count -gt 0) {
    Write-Host "`n2. Getting next item..." -ForegroundColor Yellow
    $item = Get-NextQueueItem

    if ($item) {
        Write-Host "Got item: $($item.FileName)" -ForegroundColor Green

        Write-Host "`n3. Processing document..." -ForegroundColor Yellow

        # Simulate processing (since we're in SkipPythonCheck mode)
        $outputDir = Join-Path ".\ProcessedDocuments" $item.Id
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)
        $outputFile = Join-Path $outputDir "$baseName.md"

        "Simulated conversion of: $($item.FileName)`nGenerated at: $(Get-Date)" | Set-Content $outputFile -Encoding UTF8

        Write-Host "  Created: $outputFile" -ForegroundColor Green

        # Update status
        Update-ItemStatus $item.Id @{
            Status = 'Completed'
            Progress = 100
            OutputFile = $outputFile
            EndTime = Get-Date
        }

        Write-Host "  Status updated to Completed" -ForegroundColor Green

    } else {
        Write-Host "Could not get next item (might be processing by background service)" -ForegroundColor Yellow
    }
} else {
    Write-Host "`nNo items in queue to process" -ForegroundColor Yellow
}

Write-Host "`n=== Manual Processing Complete ===" -ForegroundColor Cyan
