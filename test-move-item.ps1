Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Testing Move-Item behavior..." -ForegroundColor Cyan

$queueFile = "$env:TEMP\docling_queue.json"
$tempFile = "$queueFile.tmp"

Write-Host "Queue file: $queueFile" -ForegroundColor Gray
Write-Host "Temp file: $tempFile" -ForegroundColor Gray

# Create test content
$testItem = @{
    Id = "test-123"
    FileName = "test.txt"
}

# Write to temp file
Write-Host "`nWriting to temp file..." -ForegroundColor Yellow
$content = "[" + ($testItem | ConvertTo-Json -Compress) + "]"
$content | Set-Content $tempFile -Encoding UTF8

# Check temp file
if (Test-Path $tempFile) {
    $tempSize = (Get-Item $tempFile).Length
    $tempContent = Get-Content $tempFile -Raw
    Write-Host "Temp file created successfully" -ForegroundColor Green
    Write-Host "  Size: $tempSize bytes" -ForegroundColor White
    Write-Host "  Content: $tempContent" -ForegroundColor White
} else {
    Write-Host "ERROR: Temp file was NOT created!" -ForegroundColor Red
    exit 1
}

# Move to queue file
Write-Host "`nMoving temp file to queue file..." -ForegroundColor Yellow
try {
    Move-Item -Path $tempFile -Destination $queueFile -Force -ErrorAction Stop
    Write-Host "Move-Item succeeded" -ForegroundColor Green
} catch {
    Write-Host "ERROR during Move-Item: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Check queue file IMMEDIATELY
Write-Host "`nChecking queue file IMMEDIATELY after move..." -ForegroundColor Yellow
if (Test-Path $queueFile) {
    $queueSize = (Get-Item $queueFile).Length
    $queueContent = Get-Content $queueFile -Raw
    Write-Host "Queue file exists" -ForegroundColor Green
    Write-Host "  Size: $queueSize bytes" -ForegroundColor White
    Write-Host "  Content: $queueContent" -ForegroundColor White

    if ($queueContent.Trim() -eq "[]") {
        Write-Host "`nWARNING: File was moved but content is empty []!" -ForegroundColor Red
        Write-Host "This suggests something is overwriting the file" -ForegroundColor Red
    } else {
        Write-Host "`nSUCCESS: File contains expected content!" -ForegroundColor Green
    }
} else {
    Write-Host "ERROR: Queue file does NOT exist after move!" -ForegroundColor Red
}
