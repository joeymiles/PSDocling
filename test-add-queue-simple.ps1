Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Testing Add-QueueItem..." -ForegroundColor Cyan

# Clear queue first
"[]" | Set-Content "$env:TEMP\docling_queue.json" -Encoding UTF8

# Create test item
$testItem = @{
    Id = [guid]::NewGuid().ToString()
    FileName = "test.html"
    FilePath = "C:\test.html"
    Status = "Queued"
}

Write-Host "`nAdding item to queue..." -ForegroundColor Yellow
Add-QueueItem $testItem

Write-Host "Waiting 2 seconds..." -ForegroundColor Gray
Start-Sleep 2

Write-Host "`nChecking queue..." -ForegroundColor Yellow
$queueContent = Get-Content "$env:TEMP\docling_queue.json" -Raw
Write-Host "Queue content: $queueContent" -ForegroundColor White

$items = Get-QueueItems
Write-Host "Items in queue: $($items.Count)" -ForegroundColor $(if ($items.Count -gt 0) { 'Green' } else { 'Red' })

if ($items.Count -gt 0) {
    Write-Host "`nSUCCESS: Add-QueueItem is working!" -ForegroundColor Green
    Write-Host "Item ID: $($items[0].Id)" -ForegroundColor White
} else {
    Write-Host "`nFAILURE: Queue is empty!" -ForegroundColor Red
}
