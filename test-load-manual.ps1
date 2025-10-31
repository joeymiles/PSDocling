# Manual load test - Add to queue and manually trigger
Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "=== PSDocling Load Test (Manual Trigger) ===" -ForegroundColor Cyan

# Get the documents that are in Ready status from previous test
$status = Get-ProcessingStatus
$readyDocs = $status.GetEnumerator() | Where-Object { $_.Value.Status -eq 'Ready' } | Select-Object -First 3

if ($readyDocs.Count -eq 0) {
    Write-Host "No documents in Ready status. Queueing new documents..." -ForegroundColor Yellow
    $testPath = (Resolve-Path 'test.html').Path
    $docIds = @()
    for ($i = 1; $i -le 3; $i++) {
        $docId = Add-DocumentToQueue -Path $testPath -ExportFormat 'markdown'
        $docIds += $docId
        Write-Host "Queued document $i with ID: $docId" -ForegroundColor Green
    }
} else {
    $docIds = $readyDocs | ForEach-Object { $_.Key }
    Write-Host "Found $($docIds.Count) documents in Ready status" -ForegroundColor Green
    $docIds | ForEach-Object {
        Write-Host "  - $_" -ForegroundColor Gray
    }
}

Write-Host "`nManually triggering conversions for $($docIds.Count) documents..." -ForegroundColor Cyan
foreach ($docId in $docIds) {
    Write-Host "Starting conversion: $($docId.Substring(0, 8))..." -ForegroundColor Gray
    try {
        $result = Start-DocumentConversion -DocumentId $docId -ExportFormat 'markdown'
        if ($result) {
            Write-Host "  -> Started successfully" -ForegroundColor Green
        } else {
            Write-Host "  -> Failed to start" -ForegroundColor Red
        }
    } catch {
        Write-Host "  -> Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nMonitoring processing (checking every 2 seconds for up to 30 seconds)..." -ForegroundColor Yellow

$maxWait = 30
$elapsed = 0
$allCompleted = $false

while ($elapsed -lt $maxWait -and -not $allCompleted) {
    Start-Sleep 2
    $elapsed += 2

    $status = Get-ProcessingStatus
    $completed = 0
    $processing = 0
    $ready = 0
    $errors = 0

    foreach ($docId in $docIds) {
        if ($status[$docId]) {
            switch ($status[$docId].Status) {
                'Completed' { $completed++ }
                'Processing' { $processing++ }
                'Ready' { $ready++ }
                'Error' { $errors++ }
            }
        }
    }

    $statusColor = if ($completed -eq $docIds.Count) { 'Green' } elseif ($processing -gt 0 -or $completed -gt 0) { 'Yellow' } else { 'Cyan' }
    Write-Host "[${elapsed}s] Completed=$completed, Processing=$processing, Ready=$ready, Errors=$errors" -ForegroundColor $statusColor

    if ($completed -eq $docIds.Count) {
        $allCompleted = $true
        Write-Host "`nAll documents completed!" -ForegroundColor Green
        break
    }
}

Write-Host "`n=== Final Results ===" -ForegroundColor Cyan
$status = Get-ProcessingStatus
$results = @()

foreach ($docId in $docIds) {
    if ($status[$docId]) {
        $doc = $status[$docId]
        $results += [PSCustomObject]@{
            DocumentID = $docId.Substring(0, 8) + "..."
            Status = $doc.Status
            Progress = if ($doc.Progress) { "$($doc.Progress)%" } else { "0%" }
            OutputExists = if ($doc.OutputFile -and (Test-Path $doc.OutputFile)) { "YES" } else { "NO" }
        }
    }
}

$results | Format-Table -AutoSize

$completedCount = ($results | Where-Object { $_.Status -eq 'Completed' }).Count
Write-Host "`nCompleted: $completedCount / $($docIds.Count)" -ForegroundColor $(if ($completedCount -eq $docIds.Count) { 'Green' } else { 'Yellow' })

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan
