Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Testing Add-QueueItem without .GetNewClosure()..." -ForegroundColor Cyan

# Manually implement what Add-QueueItem does, but simpler
$queueFile = "$env:TEMP\docling_queue.json"
Write-Host "Queue file: $queueFile" -ForegroundColor Gray

# Create a test item
$testItem = @{
    Id = "test-123"
    FileName = "test.txt"
    Status = "Queued"
}

Write-Host "`nManually adding to queue (no mutex, no closure)..." -ForegroundColor Yellow

# Read current queue
$queue = @()
if (Test-Path $queueFile) {
    $content = Get-Content $queueFile -Raw
    if ($content.Trim() -ne "[]") {
        $queue = @($content | ConvertFrom-Json)
    }
}

Write-Host "Current queue has $($queue.Count) items" -ForegroundColor White

# Add new item
$newQueue = @($queue) + @($testItem)
Write-Host "New queue will have $($newQueue.Count) items" -ForegroundColor White

# Write back
$tempFile = "$queueFile.tmp"
if ($newQueue.Count -eq 1) {
    "[" + ($newQueue[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
} else {
    $newQueue | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
}
Move-Item -Path $tempFile -Destination $queueFile -Force

Write-Host "Wrote to queue file" -ForegroundColor Green

# Verify
Start-Sleep 1
$verifyContent = Get-Content $queueFile -Raw
Write-Host "`nVerification:" -ForegroundColor Cyan
Write-Host "File size: $((Get-Item $queueFile).Length) bytes" -ForegroundColor White
Write-Host "Content: $verifyContent" -ForegroundColor Gray

$verifyQueue = Get-QueueItems
Write-Host "`nQueue items via Get-QueueItems: $($verifyQueue.Count)" -ForegroundColor $(if ($verifyQueue.Count -gt 0) { 'Green' } else { 'Red' })
