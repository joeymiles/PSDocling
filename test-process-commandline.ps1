# Test if CommandLine property is available
Write-Host "Testing CommandLine property availability..." -ForegroundColor Cyan

$proc = Get-Process powershell -ErrorAction SilentlyContinue | Select-Object -First 1

Write-Host "`nProcess properties:" -ForegroundColor Yellow
$proc | Get-Member -MemberType Property | Select-Object Name | Format-Wide -Column 3

Write-Host "`nChecking for CommandLine property:" -ForegroundColor Yellow
if ($proc.PSObject.Properties['CommandLine']) {
    Write-Host "CommandLine property EXISTS" -ForegroundColor Green
    Write-Host "Value: $($proc.CommandLine)" -ForegroundColor White
} else {
    Write-Host "CommandLine property DOES NOT EXIST (this is the bug!)" -ForegroundColor Red
    Write-Host "This is why Stop-All.ps1 cannot find Docling processes" -ForegroundColor Yellow
}

Write-Host "`nAlternative: Get CommandLine via WMI" -ForegroundColor Yellow
$wmiProc = Get-WmiObject Win32_Process -Filter "ProcessId = $($proc.Id)"
if ($wmiProc) {
    Write-Host "WMI CommandLine: $($wmiProc.CommandLine)" -ForegroundColor Green
} else {
    Write-Host "WMI query failed" -ForegroundColor Red
}
