Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Detailed tracing of queue file modifications..." -ForegroundColor Cyan

$queueFile = "$env:TEMP\docling_queue.json"
$testItem = @{ Id = "test-123"; FileName = "test.txt"; Status = "Queued" }

function Check-QueueFile {
    param([string]$Step)
    if (Test-Path $queueFile) {
        $size = (Get-Item $queueFile).Length
        $content = Get-Content $queueFile -Raw
        Write-Host "[$Step] Size: $size bytes, Content preview: $($content.Substring(0, [Math]::Min(50, $content.Length)))" -ForegroundColor Gray
    } else {
        Write-Host "[$Step] File does NOT exist" -ForegroundColor Red
    }
}

Check-QueueFile "Initial"

# Write directly without any functions
Write-Host "`n1. Writing to temp file..." -ForegroundColor Yellow
$tempFile = "$queueFile.tmp"
"[" + ($testItem | ConvertTo-Json -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
Check-QueueFile "After temp write"

Write-Host "`n2. Moving temp to queue..." -ForegroundColor Yellow
Move-Item -Path $tempFile -Destination $queueFile -Force
Check-QueueFile "Immediately after move"

Write-Host "`n3. Sleeping..." -ForegroundColor Yellow
Start-Sleep -Milliseconds 100
Check-QueueFile "After 100ms sleep"

Write-Host "`n4. Calling Get-Content..." -ForegroundColor Yellow
$readContent = Get-Content $queueFile -Raw
Write-Host "Read: $readContent" -ForegroundColor White
Check-QueueFile "After Get-Content"

Write-Host "`n5. Calling module function Get-QueueItems..." -ForegroundColor Yellow
$items = Get-QueueItems
Write-Host "Get-QueueItems returned $($items.Count) items" -ForegroundColor White
Check-QueueFile "After Get-QueueItems"

Write-Host "`nTest complete" -ForegroundColor Cyan
