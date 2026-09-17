#Requires -Version 5.1
<#
.SYNOPSIS
    Lifecycle checks for the local desktop app standard (#44).
.DESCRIPTION
    Builds the module to a temp folder and runs it with an isolated
    PSDOCLING_HOME on a spare port, then verifies:
      1. POST /api/shutdown (Quit) stops the API and processor and frees the port
      2. Closing the pywebview window stops everything   (-WithWindow; opens a real window briefly)
      3. The API stops itself after the idle timeout
    Needs no Docling conversion. Exit 0 on success, 1 on failure.
#>
param(
    [int]$Port = 8093,
    [switch]$WithWindow
)
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$failed = 0

function Assert-True([bool]$Condition, [string]$Message) {
    if ($Condition) { Write-Host "PASS: $Message" -ForegroundColor Green }
    else { Write-Host "FAIL: $Message" -ForegroundColor Red; $script:failed++ }
}

function Test-Listening([int]$P) {
    @(Get-NetTCPConnection -LocalPort $P -State Listen -ErrorAction SilentlyContinue).Count -gt 0
}

function Wait-Until([scriptblock]$Condition, [int]$Seconds) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        if (& $Condition) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return [bool](& $Condition)
}

function Test-Alive([int[]]$Ids) {
    @($Ids | Where-Object { $_ -and (Get-Process -Id $_ -ErrorAction SilentlyContinue) }).Count -gt 0
}

if (Test-Listening $Port) { throw "Port $Port is already in use; pass -Port" }

$work = Join-Path $env:TEMP ("PSDoclingLifecycle_" + [guid]::NewGuid().ToString('N'))
$build = Join-Path $work 'build'
$env:PSDOCLING_HOME = Join-Path $work 'home'
& (Join-Path $repoRoot 'scripts\Build-PSDoclingModule.ps1') -OutputPath $build *> $null
Copy-Item (Join-Path $repoRoot 'DoclingFrontend') (Join-Path $build 'DoclingFrontend') -Recurse
Copy-Item (Join-Path $repoRoot 'scripts\Launch-PyWebView.py') $build
Import-Module (Join-Path $build 'PSDocling.psm1') -Force *> $null
Initialize-DoclingSystem -SkipPythonCheck *> $null
$runDir = Join-Path $env:PSDOCLING_HOME 'run'
$tokenFile = Join-Path $runDir 'token.txt'

try {
    # 1. Quit via API
    $sys = Start-DoclingSystem -Port $Port 3> $null 6> $null
    Assert-True $sys.Ready "Stack starts on port $Port"
    $ids = @($sys.API.Id, $sys.Processor.Id)
    $token = (Get-Content $tokenFile -Raw).Trim()
    $resp = Invoke-RestMethod "http://localhost:$Port/api/shutdown" -Method POST -Headers @{ 'X-PSDocling-Token' = $token } -TimeoutSec 5
    Assert-True ($resp.success -eq $true) "Shutdown endpoint answers before stopping"
    Assert-True (Wait-Until { -not (Test-Listening $Port) } 15) "Port $Port released after Quit"
    Assert-True (Wait-Until { -not (Test-Alive $ids) } 15) "API and processor processes exited after Quit"
    Assert-True (-not (Test-Path $tokenFile)) "Token file removed on shutdown"

    # 2. Closing the window
    if ($WithWindow) {
        $sys = Start-DoclingSystem -Port $Port -UseWebView 3> $null 6> $null
        Assert-True ($null -ne $sys.PyWebView) "pywebview window started"
        if ($sys.PyWebView) {
            $win = $sys.PyWebView
            $ids = @($sys.API.Id, $sys.Processor.Id, $win.Id)
            Assert-True (Wait-Until { $win.Refresh(); $win.MainWindowHandle -ne 0 } 30) "Window is shown"
            [void]$win.CloseMainWindow()
            Assert-True (Wait-Until { -not (Test-Listening $Port) } 20) "Port released after closing the window"
            Assert-True (Wait-Until { -not (Test-Alive $ids) } 20) "All processes exited after closing the window"
        }
    }

    # 3. Idle shutdown (API alone, short timeout)
    New-Item -ItemType Directory -Path $runDir -Force | Out-Null
    $script = Join-Path $runDir 'idle_api.ps1'
    $modulePath = (Join-Path $build 'PSDocling.psm1') -replace "'", "''"
    "Import-Module '$modulePath' -Force *> `$null; Start-APIServer -Port $Port -IdleShutdownSeconds 4" | Set-Content $script -Encoding UTF8
    $api = Start-Process powershell -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$script`"" -PassThru -WindowStyle Hidden
    Assert-True (Wait-Until { Test-Listening $Port } 20) "Idle test API listening"
    Assert-True (Wait-Until { $api.HasExited } 20) "API exits by itself after idle timeout"
    Assert-True (-not (Test-Listening $Port)) "Port released after idle shutdown"
}
finally {
    Stop-DoclingSystem *> $null
    Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\PSDOCLING_HOME -ErrorAction SilentlyContinue
}

Write-Host ""
if ($failed -gt 0) { Write-Host "Lifecycle: $failed failed" -ForegroundColor Red; exit 1 }
Write-Host "Lifecycle: all passed" -ForegroundColor Green
exit 0
