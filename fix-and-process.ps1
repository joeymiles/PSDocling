# Fix the queue and manually process the documents
Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "=== Manual Processing Workaround ===" -ForegroundColor Cyan
Write-Host "Due to queue bug, manually processing documents`n" -ForegroundColor Yellow

# Get the queued documents
$status = Get-ProcessingStatus
$queuedDocs = $status.GetEnumerator() | Where-Object { $_.Value.Status -in @('Queued', 'Ready') } | Select-Object -First 3

Write-Host "Found $($queuedDocs.Count) documents to process" -ForegroundColor Green

foreach ($doc in $queuedDocs) {
    $docId = $doc.Key
    $docInfo = $doc.Value

    Write-Host "`nProcessing: $($docInfo.FileName)" -ForegroundColor Yellow
    Write-Host "  ID: $($docId.Substring(0, 13))..." -ForegroundColor Gray

    # Manually add to queue
    $queueItem = @{
        Id = $docId
        FilePath = $docInfo.FilePath
        FileName = $docInfo.FileName
        ExportFormat = if ($docInfo.ExportFormat) { $docInfo.ExportFormat } else { 'markdown' }
        Status = 'Queued'
        QueuedTime = Get-Date
    }

    Add-QueueItem $queueItem
    Write-Host "  Added to queue file" -ForegroundColor Green
}

Write-Host "`nWaiting for processor to pick up items (30 seconds)..." -ForegroundColor Cyan

$maxWait = 30
$elapsed = 0

while ($elapsed -lt $maxWait) {
    Start-Sleep 3
    $elapsed += 3

    $status = Get-ProcessingStatus
    $completed = 0
    $processing = 0

    foreach ($doc in $queuedDocs) {
        $docId = $doc.Key
        if ($status[$docId]) {
            if ($status[$docId].Status -eq 'Completed') { $completed++ }
            if ($status[$docId].Status -eq 'Processing') { $processing++ }
        }
    }

    Write-Host "  ${elapsed}s: Completed=$completed, Processing=$processing" -ForegroundColor $(if ($completed -gt 0) { 'Green' } elseif ($processing -gt 0) { 'Yellow' } else { 'Cyan' })

    if ($completed -eq 3) {
        Write-Host "`nAll 3 documents completed!" -ForegroundColor Green
        break
    }
}

Write-Host "`n=== Final Results ===" -ForegroundColor Cyan
$status = Get-ProcessingStatus

foreach ($doc in $queuedDocs) {
    $docId = $doc.Key
    $docStatus = $status[$docId]

    Write-Host "`nDocument: $($docStatus.FileName)" -ForegroundColor White
    Write-Host "  Status: $($docStatus.Status)" -ForegroundColor $(if ($docStatus.Status -eq 'Completed') { 'Green' } else { 'Yellow' })
    Write-Host "  Progress: $($docStatus.Progress)%"

    if ($docStatus.OutputFile) {
        if (Test-Path $docStatus.OutputFile) {
            $size = [math]::Round((Get-Item $docStatus.OutputFile).Length / 1KB, 2)
            Write-Host "  Output: $($docStatus.OutputFile) ($size KB)" -ForegroundColor Green
        } else {
            Write-Host "  Output: $($docStatus.OutputFile) [MISSING]" -ForegroundColor Red
        }
    }
}

Write-Host "`n=== Workaround Complete ===" -ForegroundColor Cyan
