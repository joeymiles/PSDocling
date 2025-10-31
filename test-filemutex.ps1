Import-Module .\Build\PSDocling.psm1 -Force

Write-Host "Testing Use-FileMutex..." -ForegroundColor Cyan

$testFile = "$env:TEMP\test_mutex.txt"
Write-Host "Test file: $testFile" -ForegroundColor Gray

# Remove test file if it exists
if (Test-Path $testFile) {
    Remove-Item $testFile
}

# Test 1: Simple write
Write-Host "`nTest 1: Simple write with Use-FileMutex" -ForegroundColor Yellow
$localFile = $testFile
$result = Use-FileMutex -Name "test" -Script {
    Write-Host "Inside script block!" -ForegroundColor Magenta
    "Test content from script block" | Set-Content $localFile
    Write-Host "Wrote to file: $localFile" -ForegroundColor Magenta
}.GetNewClosure()

Start-Sleep 1

if (Test-Path $testFile) {
    $content = Get-Content $testFile -Raw
    Write-Host "SUCCESS! File was written:" -ForegroundColor Green
    Write-Host "Content: $content" -ForegroundColor White
} else {
    Write-Host "FAILURE! File was NOT written" -ForegroundColor Red
}

Write-Host "`nTest complete"
