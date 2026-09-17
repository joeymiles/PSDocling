#Requires -Version 5.1
<#
.SYNOPSIS
    Starts PSDocling as a desktop app (hidden console, own window)
.DESCRIPTION
    Used by the desktop shortcut and the root PSDocling.cmd.
      - Uses the module next to this script (installed layout), otherwise
        builds app\Build from app\Source when missing or out of date.
      - If PSDocling is already running for this data folder, only opens a
        window onto it.
      - Otherwise starts the API and processor hidden and opens the window.
    Failures are written to <home>\logs\last-launch-error.txt and shown in a
    message box, because this script normally runs without a console.
.PARAMETER Port
    Loopback port for the app (default 8080).
.PARAMETER Browser
    Open a normal browser tab instead of the app window (developer use).
#>
param(
    [int]$Port = 8080,
    [switch]$Browser
)
$ErrorActionPreference = 'Stop'

function Show-LaunchError([string]$Message) {
    try {
        $doclingHome = if ($env:PSDOCLING_HOME) { $env:PSDOCLING_HOME } else { Join-Path $env:LOCALAPPDATA 'PSDocling' }
        $logDir = Join-Path $doclingHome 'logs'
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        Set-Content -Path (Join-Path $logDir 'last-launch-error.txt') -Value "$(Get-Date -Format s) $Message" -Encoding UTF8
        Add-Content -Path (Join-Path $logDir 'launcher.log') -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] $Message" -Encoding UTF8
    } catch { }
    if ([Environment]::UserInteractive -and -not $env:PSDOCLING_NO_DIALOG) {
        Add-Type -AssemblyName System.Windows.Forms
        [void][System.Windows.Forms.MessageBox]::Show("PSDocling could not start.`n`n$Message", 'PSDocling', 'OK', 'Error')
    }
}

function Resolve-DoclingModule {
    # Installed layout: this script sits next to PSDocling.psm1
    $installed = Join-Path $PSScriptRoot 'PSDocling.psm1'
    if (Test-Path $installed) { return $installed }

    # Repo layout: app\scripts\ -> app\Build\PSDocling.psm1, rebuilt when stale
    $appDir = Split-Path $PSScriptRoot -Parent
    $source = Join-Path $appDir 'Source'
    if (-not (Test-Path $source)) { throw "PSDocling module not found next to $PSScriptRoot" }
    $built = Join-Path $appDir 'Build\PSDocling.psm1'
    $newest = Get-ChildItem $source -Recurse -Filter *.ps1 | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not (Test-Path $built) -or (Get-Item $built).LastWriteTime -lt $newest.LastWriteTime -or
        (Get-Item $built).LastWriteTime -lt (Get-Item (Join-Path $appDir 'PSDocling.psd1')).LastWriteTime) {
        & (Join-Path $PSScriptRoot 'Build-PSDoclingModule.ps1') -OutputPath (Join-Path $appDir 'Build') *> $null
    }
    if (-not (Test-Path $built)) { throw "Build did not produce $built" }
    return $built
}

try {
    $modulePath = Resolve-DoclingModule
    Import-Module $modulePath -Force *> $null
    $module = Get-Module PSDocling
    if (-not $module) { throw "Could not import $modulePath" }

    # Already running for this data folder? Just open another window onto it.
    $running = & $module {
        param($P)
        $pidFile = Join-Path (Get-DoclingPath Run) 'docling_pids.json'
        if (-not (Test-Path $pidFile)) { return $null }
        try {
            $pids = Get-Content $pidFile -Raw | ConvertFrom-Json
            $port = if ($pids.Port) { [int]$pids.Port } else { $P }
            if (-not (Get-Process -Id $pids.API -ErrorAction SilentlyContinue)) { return $null }
            Invoke-RestMethod "http://localhost:$port/api/health" -TimeoutSec 3 | Out-Null
            return $port
        } catch {
            return $null
        }
    } $Port

    if ($running) {
        & $module {
            param($P, $UseBrowser)
            Write-DoclingLog -Component launcher -Message "Already running on port $P; opening a window"
            if ($UseBrowser) { Start-Process "http://localhost:$P"; return }
            if (-not (Start-DoclingWindow -Url "http://localhost:$P" -ApiPort $P)) {
                [void](Start-DoclingAppWindow -Url "http://localhost:$P")
            }
        } $running $Browser.IsPresent *> $null
        exit 0
    }

    Initialize-DoclingSystem *> $null
    $startArgs = @{ Port = $Port }
    if ($Browser) { $startArgs.OpenBrowser = $true } else { $startArgs.UseWebView = $true }
    $result = Start-DoclingSystem @startArgs 3> $null 6> $null
    if (-not $result.Ready) {
        Stop-DoclingSystem *> $null
        throw "The PSDocling service did not start on port $Port. The port may be in use by another program."
    }
    exit 0
}
catch {
    Show-LaunchError $_.Exception.Message
    exit 1
}
