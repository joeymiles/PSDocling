# Test script to process test.html
Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Queuing test.html for processing..." -ForegroundColor Cyan
$testPath = (Resolve-Path 'test.html').Path
$docId = Add-DocumentToQueue -Path $testPath -ExportFormat 'markdown'
Write-Host "Document queued with ID: $docId" -ForegroundColor Green

Write-Host "`nWaiting for processing..." -ForegroundColor Yellow
Start-Sleep 5

Write-Host "`nChecking processing status..." -ForegroundColor Cyan
$status = Get-ProcessingStatus
if ($status) {
    $status.GetEnumerator() | ForEach-Object {
        Write-Host "`nDocument ID: $($_.Key)" -ForegroundColor Yellow
        Write-Host "Status: $($_.Value.Status)" -ForegroundColor $(if ($_.Value.Status -eq 'Completed') { 'Green' } elseif ($_.Value.Status -eq 'Error') { 'Red' } else { 'Yellow' })
        Write-Host "File: $($_.Value.FileName)"
        if ($_.Value.OutputFile) {
            Write-Host "Output: $($_.Value.OutputFile)" -ForegroundColor Green
        }
        if ($_.Value.Error) {
            Write-Host "Error: $($_.Value.Error)" -ForegroundColor Red
        }
    }
} else {
    Write-Host "No documents in processing queue" -ForegroundColor Yellow
}

Write-Host "`nQueue items:" -ForegroundColor Cyan
$queue = Get-QueueItems
if ($queue) {
    $queue | Format-Table -AutoSize
} else {
    Write-Host "Queue is empty" -ForegroundColor Yellow
}
