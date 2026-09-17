<#
.SYNOPSIS
    Developer start: runs PSDocling from this checkout with a visible console
.DESCRIPTION
    Builds app\Build from app\Source, imports that build (never an installed
    copy, so you test what you edit) and starts the API and processor.
    End users should use the desktop shortcut or PSDocling.cmd instead.
.EXAMPLE
    .\app\scripts\Start-All.ps1 -OpenBrowser
.EXAMPLE
    .\app\scripts\Start-All.ps1 -UseWebView -Port 9080
#>
param(
    [int]$Port = 8080,
    [switch]$OpenBrowser,
    [switch]$UseWebView,
    [switch]$SkipPythonCheck,
    [switch]$ClearHistory
)
$ErrorActionPreference = 'Stop'

$appDir = Split-Path -Parent $PSScriptRoot
$buildDir = Join-Path $appDir 'Build'
& (Join-Path $PSScriptRoot 'Build-PSDoclingModule.ps1') -OutputPath $buildDir | Out-Null
Import-Module (Join-Path $buildDir 'PSDocling.psm1') -Force

$initParams = @{}
if ($SkipPythonCheck) { $initParams.SkipPythonCheck = $true }
if ($ClearHistory) { $initParams.ClearHistory = $true }
Initialize-DoclingSystem @initParams | Out-Null

$startParams = @{ Port = $Port }
if ($OpenBrowser) { $startParams.OpenBrowser = $true }
if ($UseWebView) { $startParams.UseWebView = $true }
$result = Start-DoclingSystem @startParams

if ($result.Ready) {
    Write-Host "PSDocling running at $($result.Url)  (stop with .\app\scripts\Stop-All.ps1 or Quit in the app)" -ForegroundColor Green
} else {
    Write-Host "PSDocling did not start. See $env:LOCALAPPDATA\PSDocling\logs" -ForegroundColor Red
    exit 1
}
