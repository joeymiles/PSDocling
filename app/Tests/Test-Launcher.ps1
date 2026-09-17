#Requires -Version 5.1
<#
.SYNOPSIS
    Desktop launcher checks (#44). Opens a real PSDocling window briefly.
.DESCRIPTION
    Runs app\scripts\Start-PSDocling.ps1 the way the shortcut does (hidden
    powershell) with an isolated PSDOCLING_HOME and PSDOCLING_DESKTOP:
      1. launch starts API + processor + window, no visible console
      2. a second launch reuses the running API
      3. the shortcut API creates PSDocling.lnk with hidden PowerShell and the .ico
      4. closing the window stops every process and frees the port
      5. a port conflict writes last-launch-error.txt and exits non-zero
#>
param([int]$Port = 8093)
$ErrorActionPreference = 'Stop'
$appDir = Split-Path -Parent $PSScriptRoot
$launcher = Join-Path $appDir 'scripts\Start-PSDocling.ps1'
$failed = 0

function Assert-True([bool]$Condition, [string]$Message) {
    if ($Condition) { Write-Host "PASS: $Message" -ForegroundColor Green }
    else { Write-Host "FAIL: $Message" -ForegroundColor Red; $script:failed++ }
}
function Test-Listening([int]$P) { @(Get-NetTCPConnection -LocalPort $P -State Listen -ErrorAction SilentlyContinue).Count -gt 0 }
function Wait-Until([scriptblock]$Condition, [int]$Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) { if (& $Condition) { return $true }; Start-Sleep -Milliseconds 500 }
    return [bool](& $Condition)
}
function Get-Pids { Get-Content (Join-Path $env:PSDOCLING_HOME 'run\docling_pids.json') -Raw | ConvertFrom-Json }
function Invoke-Launcher([int]$LaunchPort) {
    $proc = Start-Process powershell -ArgumentList '-NoProfile', '-WindowStyle', 'Hidden', '-ExecutionPolicy', 'Bypass', '-File', "`"$launcher`"", '-Port', $LaunchPort -PassThru -WindowStyle Hidden
    $proc.WaitForExit(240000) | Out-Null
    return $proc.ExitCode
}

if (Test-Listening $Port) { throw "Port $Port is already in use; pass -Port" }
$work = Join-Path $env:TEMP ("PSDoclingLauncher_" + [guid]::NewGuid().ToString('N'))
$env:PSDOCLING_HOME = Join-Path $work 'home'
$env:PSDOCLING_DESKTOP = Join-Path $work 'desktop'
$env:PSDOCLING_NO_DIALOG = '1'
New-Item -ItemType Directory -Path $env:PSDOCLING_DESKTOP -Force | Out-Null
$blocker = $null

try {
    # 1. First launch
    $code = Invoke-Launcher $Port
    Assert-True ($code -eq 0) "Launcher exits 0 (got $code)"
    Assert-True (Test-Listening $Port) "App listening on $Port"
    $pids = Get-Pids
    Assert-True ([bool](Get-Process -Id $pids.API -ErrorAction SilentlyContinue)) "API process running"
    Assert-True ($pids.PyWebView -and (Get-Process -Id $pids.PyWebView -ErrorAction SilentlyContinue)) "Window process running"
    $win = Get-Process -Id $pids.PyWebView -ErrorAction SilentlyContinue
    Assert-True ($win -and $win.ProcessName -eq 'pythonw') "Window runs under pythonw (no console)"
    Assert-True ((Wait-Until { $win.Refresh(); $win.MainWindowHandle -ne 0 } 30) -and $win.MainWindowTitle -eq 'PSDocling') "Window titled PSDocling"
    $apiConsole = (Get-Process -Id $pids.API).MainWindowHandle
    Assert-True ($apiConsole -eq 0) "API has no visible window"

    # 2. Second launch reuses the running API
    $code = Invoke-Launcher $Port
    $pids2 = Get-Pids
    Assert-True ($code -eq 0 -and $pids2.API -eq $pids.API) "Second launch reuses running API"
    $extraWindows = Get-CimInstance Win32_Process -Filter "Name='pythonw.exe'" | Where-Object { $_.CommandLine -like '*Launch-PyWebView.py*' -and $_.ProcessId -ne $pids.PyWebView }
    foreach ($w in $extraWindows) { Stop-Process -Id $w.ProcessId -Force -ErrorAction SilentlyContinue }
    Assert-True (@($extraWindows).Count -ge 1) "Second launch opened a window onto it"

    # 3. Shortcut
    $token = (Get-Content (Join-Path $env:PSDOCLING_HOME 'run\token.txt') -Raw).Trim()
    $info = Invoke-RestMethod "http://localhost:$Port/api/app-info"
    Assert-True (-not $info.shortcutExists -and -not $info.shortcutOffered -and $info.launcherAvailable) "app-info reports shortcut not yet offered"
    $made = Invoke-RestMethod "http://localhost:$Port/api/shortcut" -Method POST -ContentType 'application/json' -Headers @{ 'X-PSDocling-Token' = $token } -Body '{"create":true}'
    $lnkPath = Join-Path $env:PSDOCLING_DESKTOP 'PSDocling.lnk'
    Assert-True ($made.success -and (Test-Path $lnkPath)) "Shortcut created on (test) desktop"
    $lnk = (New-Object -ComObject WScript.Shell).CreateShortcut($lnkPath)
    Assert-True ($lnk.TargetPath -like '*powershell.exe' -and $lnk.Arguments -like '*-WindowStyle Hidden*' -and $lnk.Arguments -like '*Start-PSDocling.ps1*') "Shortcut runs hidden PowerShell launcher"
    $iconFile = ($lnk.IconLocation -split ',')[0]
    Assert-True ($iconFile -like '*psdocling.ico' -and (Test-Path $iconFile)) "Shortcut uses psdocling.ico ($iconFile)"
    $info = Invoke-RestMethod "http://localhost:$Port/api/app-info"
    Assert-True ($info.shortcutExists -and $info.shortcutOffered) "app-info remembers the offer"

    # 4. Closing the window stops everything
    $all = @($pids.API, $pids.Processor, $pids.PyWebView)
    [void]$win.CloseMainWindow()
    Assert-True (Wait-Until { -not (Test-Listening $Port) } 20) "Port released after closing the window"
    Assert-True (Wait-Until { @($all | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue }).Count -eq 0 } 20) "All processes exited"

    # 5. Port conflict is reported
    $blocker = New-Object System.Net.HttpListener
    $blocker.Prefixes.Add("http://localhost:$Port/")
    $blocker.Start()
    Remove-Item (Join-Path $env:PSDOCLING_HOME 'logs\last-launch-error.txt') -ErrorAction SilentlyContinue
    $code = Invoke-Launcher $Port
    $errFile = Join-Path $env:PSDOCLING_HOME 'logs\last-launch-error.txt'
    Assert-True ($code -ne 0) "Launcher exits non-zero when the port is taken (got $code)"
    Assert-True ((Test-Path $errFile) -and ((Get-Content $errFile -Raw) -match 'did not start')) "last-launch-error.txt explains the failure"
}
finally {
    if ($blocker) { $blocker.Stop(); $blocker.Close() }
    $mod = Join-Path $appDir 'Build\PSDocling.psm1'
    if (Test-Path $mod) { Import-Module $mod -Force *> $null; Stop-DoclingSystem *> $null; Remove-Module PSDocling -Force -ErrorAction SilentlyContinue }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\PSDOCLING_HOME, Env:\PSDOCLING_DESKTOP, Env:\PSDOCLING_NO_DIALOG -ErrorAction SilentlyContinue
}

Write-Host ""
if ($failed -gt 0) { Write-Host "Launcher: $failed failed" -ForegroundColor Red; exit 1 }
Write-Host "Launcher: all passed" -ForegroundColor Green
exit 0
