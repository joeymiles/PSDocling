Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Testing Add-QueueItem function..." -ForegroundColor Cyan

# Get a queued document
$status = Get-ProcessingStatus
$queuedDoc = $status.GetEnumerator() | Where-Object { $_.Value.Status -eq 'Queued' } | Select-Object -First 1

if ($queuedDoc) {
    $docInfo = $queuedDoc.Value
    $docId = $queuedDoc.Key

    Write-Host "Found queued document: $($docInfo.FileName)" -ForegroundColor Yellow
    Write-Host "Document ID: $docId" -ForegroundColor Gray

    # Create queue item
    $queueItem = @{
        Id = $docId
        FilePath = $docInfo.FilePath
        FileName = $docInfo.FileName
        ExportFormat = 'markdown'
        Status = 'Queued'
        QueuedTime = Get-Date
    }

    Write-Host "`nCalling Add-QueueItem..." -ForegroundColor Cyan
    try {
        Add-QueueItem $queueItem
        Write-Host "Add-QueueItem completed" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor Red
    }

    # Check if it was added
    Start-Sleep 1
    Write-Host "`nChecking queue..." -ForegroundColor Cyan
    $queue = Get-QueueItems
    if ($queue) {
        Write-Host "Queue has $($queue.Count) items" -ForegroundColor Green
        $queue | Format-Table Id, FileName, Status -AutoSize
    } else {
        Write-Host "Queue is still EMPTY" -ForegroundColor Red
    }

    # Check file directly
    Write-Host "`nChecking queue file directly..." -ForegroundColor Cyan
    $queuePath = "$env:TEMP\docling_queue.json"
    if (Test-Path $queuePath) {
        $content = Get-Content $queuePath -Raw
        Write-Host "Queue file size: $((Get-Item $queuePath).Length) bytes" -ForegroundColor White
        Write-Host "Content: $content" -ForegroundColor Gray
    }
} else {
    Write-Host "No queued documents found" -ForegroundColor Yellow
}
