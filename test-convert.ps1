# Test script to manually trigger conversion
Import-Module .\Build\PSDocling.psm1 -Force

$docId = "f9b89356-38c9-495f-b610-5b6e5fbf17eb"

Write-Host "Starting conversion for document: $docId" -ForegroundColor Cyan
Start-DocumentConversion -DocumentId $docId -ExportFormat 'markdown'

Write-Host "`nWaiting for processing (15 seconds)..." -ForegroundColor Yellow
Start-Sleep 15

Write-Host "`nChecking final status..." -ForegroundColor Cyan
$status = Get-ProcessingStatus
if ($status -and $status[$docId]) {
    $doc = $status[$docId]
    Write-Host "`nDocument ID: $docId" -ForegroundColor Yellow
    Write-Host "Status: $($doc.Status)" -ForegroundColor $(if ($doc.Status -eq 'Completed') { 'Green' } elseif ($doc.Status -eq 'Error') { 'Red' } else { 'Yellow' })
    Write-Host "File: $($doc.FileName)"
    Write-Host "Progress: $($doc.Progress)%"
    if ($doc.OutputFile) {
        Write-Host "Output: $($doc.OutputFile)" -ForegroundColor Green
        if (Test-Path $doc.OutputFile) {
            Write-Host "`nOutput file content (first 20 lines):" -ForegroundColor Cyan
            Get-Content $doc.OutputFile -TotalCount 20
        }
    }
    if ($doc.Error) {
        Write-Host "Error: $($doc.Error)" -ForegroundColor Red
        if ($doc.ErrorDetails) {
            Write-Host "Error Details: $($doc.ErrorDetails | ConvertTo-Json)" -ForegroundColor Red
        }
    }
} else {
    Write-Host "Document not found in status" -ForegroundColor Red
}
