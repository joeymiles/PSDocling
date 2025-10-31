Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "=== Queue Debugging ===" -ForegroundColor Cyan

Write-Host "`n1. Checking queue items..." -ForegroundColor Yellow
$queueItems = Get-QueueItems
if ($queueItems) {
    Write-Host "Found $($queueItems.Count) items in queue:" -ForegroundColor Green
    $queueItems | Format-Table Id, FileName, Status -AutoSize
} else {
    Write-Host "Queue is EMPTY" -ForegroundColor Red
}

Write-Host "`n2. Trying to get next queue item..." -ForegroundColor Yellow
$nextItem = Get-NextQueueItem
if ($nextItem) {
    Write-Host "Got next item:" -ForegroundColor Green
    $nextItem | Format-List
} else {
    Write-Host "No next item available" -ForegroundColor Red
}

Write-Host "`n3. Checking status file..." -ForegroundColor Yellow
$status = Get-ProcessingStatus
$queuedDocs = $status.GetEnumerator() | Where-Object { $_.Value.Status -in @('Queued', 'Ready') }
Write-Host "Found $($queuedDocs.Count) documents in Queued/Ready status:" -ForegroundColor $(if ($queuedDocs.Count -gt 0) { 'Yellow' } else { 'Green' })
$queuedDocs | ForEach-Object {
    Write-Host "  - $($_.Key.Substring(0,13))... $($_.Value.FileName) [$($_.Value.Status)]" -ForegroundColor Gray
}

Write-Host "`n4. Checking queue file path..." -ForegroundColor Yellow
$queuePath = "$env:TEMP\docling_queue.json"
Write-Host "Queue file: $queuePath" -ForegroundColor White
if (Test-Path $queuePath) {
    $fileInfo = Get-Item $queuePath
    Write-Host "  Size: $($fileInfo.Length) bytes" -ForegroundColor Green
    Write-Host "  Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Green
    if ($fileInfo.Length -gt 0) {
        Write-Host "  Content preview:" -ForegroundColor Cyan
        Get-Content $queuePath | Select-Object -First 10 | ForEach-Object {
            Write-Host "    $_" -ForegroundColor Gray
        }
    } else {
        Write-Host "  File is EMPTY" -ForegroundColor Red
    }
} else {
    Write-Host "  File does NOT exist" -ForegroundColor Red
}

Write-Host "`n=== Debug Complete ===" -ForegroundColor Cyan
