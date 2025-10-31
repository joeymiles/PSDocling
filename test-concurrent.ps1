# Concurrent processing test - Process 3 files simultaneously
Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "====================================" -ForegroundColor Cyan
Write-Host " PSDocling Concurrent Process Test" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Create 3 test files
Write-Host "[1/5] Creating 3 test files..." -ForegroundColor Yellow
$testFiles = @()
for ($i = 1; $i -le 3; $i++) {
    $fileName = "test-doc-$i.html"
    $content = @"
<!DOCTYPE html>
<html>
<head><title>Test Document $i</title></head>
<body>
<h1>Test Document Number $i</h1>
<p>This is test document $i for concurrent processing.</p>
<p>Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<h2>Test Content</h2>
<ul>
<li>Document ID: $i</li>
<li>Purpose: Load testing</li>
<li>Type: HTML to Markdown conversion</li>
</ul>
<h2>Sample Data</h2>
<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.</p>
<p><strong>Test Status:</strong> Ready for processing</p>
</body>
</html>
"@
    $content | Out-File -FilePath $fileName -Encoding UTF8
    $testFiles += $fileName
    Write-Host "  Created: $fileName" -ForegroundColor Green
}

Write-Host ""
Write-Host "[2/5] Queuing all 3 documents..." -ForegroundColor Yellow
$docIds = @()
foreach ($file in $testFiles) {
    $fullPath = (Resolve-Path $file).Path
    $docId = Add-DocumentToQueue -Path $fullPath -ExportFormat 'markdown'
    $docIds += $docId
    Write-Host "  Queued: $file -> $($docId.Substring(0, 13))..." -ForegroundColor Green
}

Write-Host ""
Write-Host "[3/5] Starting conversions for all 3 documents..." -ForegroundColor Yellow
foreach ($docId in $docIds) {
    try {
        Start-DocumentConversion -DocumentId $docId -ExportFormat 'markdown' | Out-Null
        Write-Host "  Started: $($docId.Substring(0, 13))..." -ForegroundColor Green
    } catch {
        Write-Host "  Failed: $($docId.Substring(0, 13))... - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "[4/5] Waiting for processing to complete..." -ForegroundColor Yellow
Write-Host "(Checking every 2 seconds, max 30 seconds)" -ForegroundColor Gray
Write-Host ""

$maxWait = 30
$checkInterval = 2
$elapsed = 0
$allCompleted = $false

# Progress tracking
$progressBar = ""
while ($elapsed -lt $maxWait -and -not $allCompleted) {
    Start-Sleep $checkInterval
    $elapsed += $checkInterval

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

    # Visual progress
    $progressBar = ""
    for ($i = 0; $i -lt $completed; $i++) { $progressBar += "[DONE] " }
    for ($i = 0; $i -lt $processing; $i++) { $progressBar += "[PROC] " }
    for ($i = 0; $i -lt $queued; $i++) { $progressBar += "[WAIT] " }
    for ($i = 0; $i -lt $errors; $i++) { $progressBar += "[ERR!] " }

    $statusColor = if ($completed -eq 3) { 'Green' } elseif ($processing -gt 0) { 'Yellow' } elseif ($errors -gt 0) { 'Red' } else { 'Cyan' }
    Write-Host "  ${elapsed}s: $progressBar (C:$completed P:$processing Q:$queued E:$errors)" -ForegroundColor $statusColor

    if ($completed -eq 3) {
        $allCompleted = $true
    }
}

Write-Host ""
if ($allCompleted) {
    Write-Host "All 3 documents completed successfully!" -ForegroundColor Green
} else {
    Write-Host "Timeout reached or processing incomplete" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "[5/5] Final Results:" -ForegroundColor Yellow
Write-Host ""

$status = Get-ProcessingStatus
$results = @()
$fileIndex = 1

foreach ($docId in $docIds) {
    if ($status[$docId]) {
        $doc = $status[$docId]

        $duration = "N/A"
        if ($doc.StartTime -and $doc.EndTime) {
            $duration = [math]::Round((New-TimeSpan -Start $doc.StartTime -End $doc.EndTime).TotalSeconds, 2)
        }

        $outputStatus = "NO"
        $outputSize = "-"
        if ($doc.OutputFile -and (Test-Path $doc.OutputFile)) {
            $outputStatus = "YES"
            $outputSize = "$([math]::Round((Get-Item $doc.OutputFile).Length / 1KB, 2)) KB"
        }

        $results += [PSCustomObject]@{
            File = "test-doc-$fileIndex.html"
            DocID = $docId.Substring(0, 8) + "..."
            Status = $doc.Status
            Progress = if ($doc.Progress) { "$($doc.Progress)%" } else { "0%" }
            Duration = "${duration}s"
            Output = $outputStatus
            Size = $outputSize
        }
        $fileIndex++
    }
}

$results | Format-Table -AutoSize

# Summary statistics
Write-Host ""
Write-Host "====================================" -ForegroundColor Cyan
Write-Host " Test Summary" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
$completedCount = ($results | Where-Object { $_.Status -eq 'Completed' }).Count
$successRate = [math]::Round(($completedCount / 3) * 100, 1)
Write-Host "Documents Processed: $completedCount / 3 ($successRate%)" -ForegroundColor $(if ($completedCount -eq 3) { 'Green' } else { 'Yellow' })
Write-Host "Total Time: ${elapsed}s" -ForegroundColor White

$totalSize = 0
foreach ($docId in $docIds) {
    if ($status[$docId] -and $status[$docId].OutputFile -and (Test-Path $status[$docId].OutputFile)) {
        $totalSize += (Get-Item $status[$docId].OutputFile).Length
    }
}
Write-Host "Total Output: $([math]::Round($totalSize / 1KB, 2)) KB" -ForegroundColor White

# Check if output files exist
Write-Host ""
Write-Host "Output Files:" -ForegroundColor Cyan
foreach ($docId in $docIds) {
    if ($status[$docId] -and $status[$docId].OutputFile) {
        $outputPath = $status[$docId].OutputFile
        if (Test-Path $outputPath) {
            Write-Host "  [OK] $outputPath" -ForegroundColor Green
        } else {
            Write-Host "  [MISSING] $outputPath" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "====================================" -ForegroundColor Cyan
Write-Host " Test Complete!" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Cyan
