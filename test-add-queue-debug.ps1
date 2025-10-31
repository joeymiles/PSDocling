Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Testing Add-QueueItem with debug..." -ForegroundColor Cyan

# Clear queue first
"[]" | Set-Content "$env:TEMP\docling_queue.json" -Encoding UTF8

# Create test item
$testItem = @{
    Id = "test-123"
    FileName = "test.html"
    FilePath = "C:\test.html"
    Status = "Queued"
}

Write-Host "`nQueue before: $(Get-Content "$env:TEMP\docling_queue.json" -Raw)" -ForegroundColor Gray

Write-Host "`nCalling Add-QueueItem..." -ForegroundColor Yellow
try {
    Add-QueueItem $testItem
    Write-Host "Add-QueueItem completed without error" -ForegroundColor Green
} catch {
    Write-Host "ERROR in Add-QueueItem: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}

Write-Host "`nQueue after: $(Get-Content "$env:TEMP\docling_queue.json" -Raw)" -ForegroundColor Gray

$items = Get-QueueItems
Write-Host "`nItems in queue: $($items.Count)" -ForegroundColor $(if ($items.Count -gt 0) { 'Green' } else { 'Red' })
