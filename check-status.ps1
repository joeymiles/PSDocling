Import-Module .\Build\PSDocling.psm1 -Force
Write-Host "Checking document status..." -ForegroundColor Cyan
$status = Get-ProcessingStatus
$status.GetEnumerator() | Where-Object { $_.Value.FileName -like 'test-doc-*.html' } | ForEach-Object {
    Write-Host "Doc: $($_.Key.Substring(0,13))... File: $($_.Value.FileName) Status: $($_.Value.Status) Progress: $($_.Value.Progress)%"
}
