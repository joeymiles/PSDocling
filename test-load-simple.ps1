# Simplified load test - Let processor pick up queued items automatically
Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "=== PSDocling Load Test (Auto-Processing) ===" -ForegroundColor Cyan
Write-Host "Queueing 3 documents for automatic processing`n" -ForegroundColor Yellow

$testPath = (Resolve-Path 'test.html').Path
$docIds = @()

# Queue 3 documents - processor should pick them up automatically
for ($i = 1; $i -le 3; $i++) {
    Write-Host "Queuing document $i..." -ForegroundColor Gray
    $docId = Add-DocumentToQueue -Path $testPath -ExportFormat 'markdown'
    $docIds += $docId
    Write-Host "  -> Queued with ID: $docId" -ForegroundColor Green
}

Write-Host "`n$($docIds.Count) documents queued. Waiting for processor to pick them up..." -ForegroundColor Cyan

# Wait a moment and check queue
Start-Sleep 2
Write-Host "`nChecking queue status..." -ForegroundColor Yellow
$queue = Get-QueueItems
if ($queue) {
    Write-Host "Queue has $($queue.Count) items waiting" -ForegroundColor Green
    $queue | ForEach-Object {
        Write-Host "  - $($_.FileName) (Status: $($_.Status))" -ForegroundColor Gray
    }
} else {
    Write-Host "Queue is empty (items may have been picked up already)" -ForegroundColor Yellow
}

Write-Host "`nMonitoring processing progress..." -ForegroundColor Yellow
Write-Host "(Checking every 3 seconds for up to 45 seconds)`n" -ForegroundColor Gray

$maxWait = 45
$elapsed = 0
$allCompleted = $false

while ($elapsed -lt $maxWait -and -not $allCompleted) {
    Start-Sleep 3
    $elapsed += 3

    $status = Get-ProcessingStatus
    $completed = 0
    $processing = 0
    $queued = 0
    $errors = 0
    $ready = 0

    foreach ($docId in $docIds) {
        if ($status[$docId]) {
            switch ($status[$docId].Status) {
                'Completed' { $completed++ }
                'Processing' { $processing++ }
                'Queued' { $queued++ }
                'Ready' { $ready++ }
                'Error' { $errors++ }
            }
        }
    }

    $statusColor = if ($completed -eq $docIds.Count) { 'Green' } elseif ($processing -gt 0) { 'Yellow' } else { 'Cyan' }
    Write-Host "[${elapsed}s] Completed=$completed, Processing=$processing, Ready=$ready, Queued=$queued, Errors=$errors" -ForegroundColor $statusColor

    if ($completed -eq $docIds.Count) {
        $allCompleted = $true
        Write-Host "`nAll documents completed successfully!" -ForegroundColor Green
    }
}

if (-not $allCompleted) {
    Write-Host "`nTimeout reached. Checking final status..." -ForegroundColor Yellow
}

Write-Host "`n=== Final Results ===" -ForegroundColor Cyan
$status = Get-ProcessingStatus

$results = @()
foreach ($docId in $docIds) {
    if ($status[$docId]) {
        $doc = $status[$docId]

        $duration = "N/A"
        if ($doc.StartTime -and $doc.EndTime) {
            $duration = [math]::Round((New-TimeSpan -Start $doc.StartTime -End $doc.EndTime).TotalSeconds, 2)
        }

        $outputExists = "NO"
        if ($doc.OutputFile -and (Test-Path $doc.OutputFile)) {
            $outputExists = "YES"
        }

        $results += [PSCustomObject]@{
            DocumentID = $docId.Substring(0, 8) + "..."
            FileName = $doc.FileName
            Status = $doc.Status
            Progress = if ($doc.Progress) { "$($doc.Progress)%" } else { "0%" }
            Duration = $duration
            OutputFile = $outputExists
        }
    }
}

$results | Format-Table -AutoSize

# Summary
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$totalCompleted = ($results | Where-Object { $_.Status -eq 'Completed' }).Count
$totalErrors = ($results | Where-Object { $_.Status -eq 'Error' }).Count
$totalProcessing = ($results | Where-Object { $_.Status -eq 'Processing' }).Count
$totalQueued = ($results | Where-Object { $_.Status -in @('Queued', 'Ready') }).Count

Write-Host "Total Documents: $($docIds.Count)" -ForegroundColor White
Write-Host "Completed: $totalCompleted" -ForegroundColor $(if ($totalCompleted -eq $docIds.Count) { 'Green' } else { 'Yellow' })
Write-Host "Processing: $totalProcessing" -ForegroundColor Yellow
Write-Host "Queued/Ready: $totalQueued" -ForegroundColor Cyan
Write-Host "Errors: $totalErrors" -ForegroundColor $(if ($totalErrors -eq 0) { 'Green' } else { 'Red' })

# Check files created
Write-Host "`n=== Output Files ===" -ForegroundColor Cyan
$outputDir = ".\ProcessedDocuments"
foreach ($docId in $docIds) {
    $docDir = Join-Path $outputDir $docId
    if (Test-Path $docDir) {
        $files = Get-ChildItem $docDir -File
        Write-Host "Document $($docId.Substring(0, 8))...: $($files.Count) file(s)" -ForegroundColor Green
        foreach ($file in $files) {
            Write-Host "  - $($file.Name) ($([math]::Round($file.Length / 1KB, 2)) KB)" -ForegroundColor Gray
        }
    } else {
        Write-Host "Document $($docId.Substring(0, 8))...: No output directory" -ForegroundColor Yellow
    }
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan
