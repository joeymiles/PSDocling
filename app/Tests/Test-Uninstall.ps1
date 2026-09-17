#Requires -Version 5.1
<#
.SYNOPSIS
    Install and uninstall checks (#44). Opens a real PSDocling window briefly.
.DESCRIPTION
    Everything is redirected to a scratch folder (module path, desktop, data
    home, TEMP), so the real installed module and user data are untouched.
    Python packages are never removed by this test.
      1. Installer copies launcher, uninstaller and icon and creates the shortcut
      2. Installed app launches from the module folder
      3. In-app Uninstall (POST /api/uninstall) stops everything and removes the
         module, shortcut, data folder and legacy TEMP files
      4. Console uninstall (-Force) leaves a foreign PSDocling.lnk alone
#>
param([int]$Port = 8093)
$ErrorActionPreference = 'Stop'
$appDir = Split-Path -Parent $PSScriptRoot
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
function Invoke-Hidden([string]$Script, [string[]]$Arguments) {
    $all = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$Script`"") + $Arguments
    $proc = Start-Process powershell -ArgumentList $all -PassThru -WindowStyle Hidden
    $proc.WaitForExit(300000) | Out-Null
    return $proc.ExitCode
}

if (Test-Listening $Port) { throw "Port $Port is already in use; pass -Port" }
$work = Join-Path $env:TEMP ("PSDoclingUninstall_" + [guid]::NewGuid().ToString('N'))
$originalTemp = $env:TEMP
$modules = Join-Path $work 'modules'
$env:PSDOCLING_HOME = Join-Path $work 'home'
$env:PSDOCLING_DESKTOP = Join-Path $work 'desktop'
$env:PSDOCLING_MODULE_BASE = $modules
$env:PSDOCLING_NO_DIALOG = '1'
$env:TEMP = Join-Path $work 'temp'
$env:TMP = $env:TEMP
New-Item -ItemType Directory -Path $env:PSDOCLING_DESKTOP, $env:TEMP, $modules -Force | Out-Null
$installed = Join-Path $modules 'PSDocling'
$shortcut = Join-Path $env:PSDOCLING_DESKTOP 'PSDocling.lnk'

try {
    # 1. Install
    $code = Invoke-Hidden (Join-Path $appDir 'scripts\Install-DoclingModule.ps1') @('-Destination', "`"$modules`"", '-DesktopShortcut', '-Force')
    Assert-True ($code -eq 0) "Installer exits 0 (got $code)"
    foreach ($f in 'PSDocling.psm1', 'Start-PSDocling.ps1', 'Uninstall-PSDocling.ps1', 'Launch-PyWebView.py', 'DoclingFrontend\psdocling.ico', 'DoclingFrontend\index.html') {
        Assert-True (Test-Path (Join-Path $installed $f)) "Installed $f"
    }
    Assert-True (Test-Path $shortcut) "Installer created the desktop shortcut"
    if (Test-Path $shortcut) {
        $lnk = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcut)
        Assert-True ($lnk.Arguments -like "*$installed\Start-PSDocling.ps1*") "Shortcut points at the installed launcher"
    }

    # 2. Launch the installed app, add some user data and legacy TEMP files
    $code = Invoke-Hidden (Join-Path $installed 'Start-PSDocling.ps1') @('-Port', $Port)
    Assert-True ($code -eq 0 -and (Test-Listening $Port)) "Installed app starts"
    $pids = Get-Content (Join-Path $env:PSDOCLING_HOME 'run\docling_pids.json') -Raw | ConvertFrom-Json
    $outDir = Join-Path $env:PSDOCLING_HOME 'data\output\doc1'
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    Set-Content (Join-Path $outDir 'result.md') 'converted'
    Set-Content (Join-Path $env:TEMP 'docling_status.json') '{}'
    New-Item -ItemType Directory -Path (Join-Path $env:TEMP 'DoclingOutput') -Force | Out-Null

    # 3. In-app uninstall
    $token = (Get-Content (Join-Path $env:PSDOCLING_HOME 'run\token.txt') -Raw).Trim()
    $resp = Invoke-RestMethod "http://localhost:$Port/api/uninstall" -Method POST -ContentType 'application/json' -Headers @{ 'X-PSDocling-Token' = $token } -Body '{"removeDoclingPackages":false,"removeTokenizers":false,"removePyWebView":false}'
    Assert-True ($resp.success -eq $true) "Uninstall endpoint accepted"
    $all = @($pids.API, $pids.Processor, $pids.PyWebView) | Where-Object { $_ }
    Assert-True (Wait-Until { -not (Test-Listening $Port) } 30) "Port released"
    Assert-True (Wait-Until { @($all | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue }).Count -eq 0 } 30) "All app processes exited"
    Assert-True (Wait-Until { -not (Test-Path $installed) } 60) "Module folder removed"
    Assert-True (Wait-Until { -not (Test-Path $env:PSDOCLING_HOME) } 30) "Data folder (with converted output) deleted"
    Assert-True (-not (Test-Path $shortcut)) "Desktop shortcut removed"
    Assert-True (-not (Test-Path (Join-Path $env:TEMP 'docling_status.json')) -and -not (Test-Path (Join-Path $env:TEMP 'DoclingOutput'))) "Legacy TEMP data removed"
    $log = Join-Path $env:TEMP 'PSDocling-uninstall.log'
    Assert-True (Wait-Until { (Test-Path $log) -and ((Get-Content $log -Raw) -match 'PSDocling uninstalled') } 30) "Uninstall log reports success"
    Assert-True (@(Get-ChildItem $env:TEMP -Filter 'PSDocling-uninstall-*.ps1').Count -eq 0) "Temporary uninstaller copy cleaned up"
    Assert-True ((Get-Content $log -Raw) -notmatch 'pip uninstall|Removing Python packages') "No Python packages removed when boxes are unchecked"

    # 4. Console uninstall keeps a foreign shortcut
    $code = Invoke-Hidden (Join-Path $appDir 'scripts\Install-DoclingModule.ps1') @('-Destination', "`"$modules`"", '-NoDesktopShortcut', '-Force')
    $foreign = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcut)
    $foreign.TargetPath = Join-Path $env:SystemRoot 'notepad.exe'
    $foreign.Save()
    $code = Invoke-Hidden (Join-Path $installed 'Uninstall-PSDocling.ps1') @('-Force')
    Assert-True ($code -eq 0) "Console uninstall exits 0 (got $code)"
    Assert-True (-not (Test-Path $installed)) "Console uninstall removed the module"
    Assert-True (Test-Path $shortcut) "Foreign PSDocling.lnk left alone"
}
finally {
    Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -and $_.CommandLine -like "*$work*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    $env:TEMP = $originalTemp
    $env:TMP = $originalTemp
    Start-Sleep -Seconds 1
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\PSDOCLING_HOME, Env:\PSDOCLING_DESKTOP, Env:\PSDOCLING_MODULE_BASE, Env:\PSDOCLING_NO_DIALOG -ErrorAction SilentlyContinue
}

Write-Host ""
if ($failed -gt 0) { Write-Host "Uninstall: $failed failed" -ForegroundColor Red; exit 1 }
Write-Host "Uninstall: all passed" -ForegroundColor Green
exit 0
