# Load test script - Queue and process multiple documents
Import-Module .\Build\PSDocling.psm1 -Force

$testFiles = @(
    @{ Path = "test.html"; Name = "Test Document 1" }
    @{ Path = "test.html"; Name = "Test Document 2" }
    @{ Path = "test.html"; Name = "Test Document 3" }
)

Write-Host "=== PSDocling Load Test ===" -ForegroundColor Cyan
Write-Host "Queueing $($testFiles.Count) documents for processing`n" -ForegroundColor Yellow

$docIds = @()
foreach ($file in $testFiles) {
    $testPath = (Resolve-Path $file.Path).Path
    Write-Host "Queuing: $($file.Name)..." -ForegroundColor Gray
    $docId = Add-DocumentToQueue -Path $testPath -ExportFormat 'markdown'
    $docIds += $docId
    Write-Host "  -> Queued with ID: $docId" -ForegroundColor Green
}

Write-Host "`n$($docIds.Count) documents queued. Starting conversions..." -ForegroundColor Cyan

# Start all conversions
foreach ($docId in $docIds) {
    Write-Host "Starting conversion: $docId" -ForegroundColor Gray
    Start-DocumentConversion -DocumentId $docId -ExportFormat 'markdown' | Out-Null
}

Write-Host "`nMonitoring processing progress..." -ForegroundColor Yellow
Write-Host "(Checking every 2 seconds for up to 30 seconds)`n" -ForegroundColor Gray

$maxWait = 30
$elapsed = 0
$allCompleted = $false

while ($elapsed -lt $maxWait -and -not $allCompleted) {
    Start-Sleep 2
    $elapsed += 2

    $status = Get-ProcessingStatus
    $completed = 0
    $processing = 0
    $queued = 0
    $errors = 0

    foreach ($docId in $docIds) {
        if ($status[$docId]) {
            switch ($status[$docId].Status) {
                'Completed' { $completed++ }
                'Processing' { $processing++ }
                'Queued' { $queued++ }
                'Error' { $errors++ }
            }
        }
    }

    Write-Host "[${elapsed}s] Status: Completed=$completed, Processing=$processing, Queued=$queued, Errors=$errors" -ForegroundColor $(if ($completed -eq $docIds.Count) { 'Green' } else { 'Yellow' })

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
        $results += [PSCustomObject]@{
            DocumentID = $docId.Substring(0, 8) + "..."
            FileName = $doc.FileName
            Status = $doc.Status
            Progress = "$($doc.Progress)%"
            Duration = if ($doc.StartTime -and $doc.EndTime) {
                [math]::Round((New-TimeSpan -Start $doc.StartTime -End $doc.EndTime).TotalSeconds, 2)
            } else {
                "N/A"
            }
            OutputFile = if ($doc.OutputFile -and (Test-Path $doc.OutputFile)) { "✓" } else { "✗" }
        }
    }
}

$results | Format-Table -AutoSize

# Summary
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
$totalCompleted = ($results | Where-Object { $_.Status -eq 'Completed' }).Count
$totalErrors = ($results | Where-Object { $_.Status -eq 'Error' }).Count
Write-Host "Total Documents: $($docIds.Count)" -ForegroundColor White
Write-Host "Completed: $totalCompleted" -ForegroundColor $(if ($totalCompleted -eq $docIds.Count) { 'Green' } else { 'Yellow' })
Write-Host "Errors: $totalErrors" -ForegroundColor $(if ($totalErrors -eq 0) { 'Green' } else { 'Red' })

# Check API status
Write-Host "`n=== API Status ===" -ForegroundColor Cyan
try {
    $apiStatus = Invoke-RestMethod -Uri "http://localhost:8080/api/status" -Method Get
    Write-Host "Queue Status:" -ForegroundColor White
    Write-Host "  Queued: $($apiStatus.QueuedCount)" -ForegroundColor Gray
    Write-Host "  Processing: $($apiStatus.ProcessingCount)" -ForegroundColor Gray
    Write-Host "  Completed: $($apiStatus.CompletedCount)" -ForegroundColor Green
    Write-Host "  Errors: $($apiStatus.ErrorCount)" -ForegroundColor $(if ($apiStatus.ErrorCount -eq 0) { 'Green' } else { 'Red' })
    Write-Host "  Total: $($apiStatus.TotalItems)" -ForegroundColor White
} catch {
    Write-Host "Could not retrieve API status: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan
