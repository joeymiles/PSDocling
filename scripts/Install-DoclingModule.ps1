#!/usr/bin/env powershell
<#
.SYNOPSIS
    Builds and installs the PSDocling PowerShell module
.DESCRIPTION
    Happy-path installer: builds from Source/ into Build/, then copies the
    module (plus DoclingFrontend and Launch-PyWebView) into the PowerShell
    modules path. No separate Build step required.
.PARAMETER Scope
    Installation scope: CurrentUser or AllUsers (default: CurrentUser)
.PARAMETER Force
    Force installation even if module already exists
.PARAMETER SkipBuild
    Skip the build step and install existing Build/ output only
.EXAMPLE
    .\scripts\Install-DoclingModule.ps1
.EXAMPLE
    .\scripts\Install-DoclingModule.ps1 -Scope AllUsers
.EXAMPLE
    .\scripts\Install-DoclingModule.ps1 -Force
#>
param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [switch]$Force,
    [switch]$SkipBuild
)

function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn($msg)  { Write-Host $msg -ForegroundColor Yellow }
function Write-Err($msg)   { Write-Host $msg -ForegroundColor Red }

function Test-IsAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

$RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $RepoRoot 'Source'))) {
    $RepoRoot = $PSScriptRoot
}

if ($Scope -eq 'AllUsers' -and -not (Test-IsAdmin)) {
    Write-Err "Installing for AllUsers requires administrator privileges. Run PowerShell as Administrator or use -Scope CurrentUser"
    exit 1
}

$moduleName = 'PSDocling'
$buildDir = Join-Path $RepoRoot 'Build'
$buildScript = Join-Path $PSScriptRoot 'Build-PSDoclingModule.ps1'

if (-not $SkipBuild) {
    if (-not (Test-Path $buildScript)) {
        Write-Err "Build script not found: $buildScript"
        exit 1
    }
    Write-Info "Building module from Source/..."
    & $buildScript -OutputPath $buildDir
    # Build-PSDoclingModule.ps1 is a PowerShell script and does not set a process
    # exit code. $LASTEXITCODE is often stale from a prior native command — do not
    # use it here. Prefer $? and verify built artifacts below.
    if (-not $?) {
        Write-Err "Build failed"
        exit 1
    }
}

$builtModule = Join-Path $buildDir 'PSDocling.psm1'
$builtManifest = Join-Path $buildDir 'PSDocling.psd1'
if (-not (Test-Path $builtModule) -or -not (Test-Path $builtManifest)) {
    Write-Err "Built module not found in $buildDir (expected PSDocling.psm1 and PSDocling.psd1)"
    Write-Info "Run .\scripts\Build-PSDoclingModule.ps1 or re-run this installer without -SkipBuild"
    exit 1
}

# Resolve install destination(s). Primary path matches the current host edition.
# When roots differ (Desktop 5.1 vs Core 7+), also install to the sibling root so
# both powershell.exe and pwsh see the same CurrentUser/AllUsers module.
$desktopUser = Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules'
$coreUser    = Join-Path $env:USERPROFILE 'Documents\PowerShell\Modules'
$desktopAll  = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
$coreAll     = Join-Path $env:ProgramFiles 'PowerShell\Modules'

$destBases = New-Object System.Collections.Generic.List[string]
if ($Scope -eq 'AllUsers') {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        [void]$destBases.Add($coreAll)
        if ($desktopAll -ne $coreAll) { [void]$destBases.Add($desktopAll) }
    } else {
        [void]$destBases.Add($desktopAll)
    }
} else {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        [void]$destBases.Add($coreUser)
        if ($desktopUser -ne $coreUser) { [void]$destBases.Add($desktopUser) }
    } else {
        [void]$destBases.Add($desktopUser)
        if ($coreUser -ne $desktopUser) { [void]$destBases.Add($coreUser) }
    }
}

$destDirs = @($destBases | ForEach-Object { Join-Path $_ $moduleName } | Select-Object -Unique)

Write-Info "Installing PSDocling module..."
Write-Info "Source (repo): $RepoRoot"
Write-Info "Build output:  $buildDir"
Write-Info "Destination(s): $($destDirs -join '; ')"
Write-Info "Scope:         $Scope"

foreach ($destDir in $destDirs) {
    if (Test-Path $destDir) {
        if ($Force) {
            Write-Warn "Module directory exists at $destDir, removing due to -Force flag..."
            Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
            Remove-Item $destDir -Recurse -Force
        } else {
            Write-Err "Module already installed at $destDir"
            Write-Info "Use -Force to overwrite, or uninstall with: .\scripts\Uninstall-DoclingModule.ps1"
            exit 1
        }
    }
}

function Copy-PSDoclingInstall {
    param([string]$DestDir)

    New-Item -Path $DestDir -ItemType Directory -Force | Out-Null

    foreach ($file in @('PSDocling.psm1', 'PSDocling.psd1', 'PSDocling.config.psd1')) {
        $sourcePath = Join-Path $buildDir $file
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath (Join-Path $DestDir $file) -Force
            Write-Info "Copied: $file (from Build) -> $DestDir"
        }
    }

    $frontendSrc = Join-Path $RepoRoot 'DoclingFrontend'
    if (Test-Path $frontendSrc) {
        Copy-Item $frontendSrc (Join-Path $DestDir 'DoclingFrontend') -Recurse -Force
        Write-Info "Copied: DoclingFrontend/ -> $DestDir"
    }

    $pyWebViewSrc = Join-Path $PSScriptRoot 'Launch-PyWebView.py'
    if (Test-Path $pyWebViewSrc) {
        Copy-Item $pyWebViewSrc (Join-Path $DestDir 'Launch-PyWebView.py') -Force
        Write-Info "Copied: Launch-PyWebView.py -> $DestDir"
    }

    foreach ($file in @('README.md', 'LICENSE', 'requirements-webview.txt')) {
        $sourcePath = Join-Path $RepoRoot $file
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath (Join-Path $DestDir $file) -Force
            Write-Info "Copied: $file -> $DestDir"
        }
    }

    $stopAll = @"
Import-Module `$PSScriptRoot -Force
Stop-DoclingSystem -ClearQueue
"@
    Set-Content -Path (Join-Path $DestDir 'Stop-All.ps1') -Value $stopAll -Encoding UTF8
}

foreach ($destDir in $destDirs) {
    Copy-PSDoclingInstall -DestDir $destDir
}

$primaryDest = $destDirs[0]
Write-Info "Testing module import from $primaryDest..."
try {
    Import-Module $primaryDest -Force
    $moduleInfo = Get-Module PSDocling
    if ($moduleInfo) {
        Write-Ok "Module installed successfully!"
        Write-Ok "Version: $($moduleInfo.Version)"
        Write-Ok "Functions exported: $($moduleInfo.ExportedFunctions.Count)"
        Write-Info ""
        Write-Info "Usage:"
        Write-Info "  Import-Module PSDocling"
        Write-Info "  Initialize-DoclingSystem -GenerateFrontend"
        Write-Info "  Start-DoclingSystem -OpenBrowser"
        Write-Info "  Stop-DoclingSystem"
        Write-Info ""
        Write-Info "Or from the repo:"
        Write-Info "  .\scripts\Start-All.ps1 -GenerateFrontend -OpenBrowser"
        Write-Info "  .\scripts\Stop-All.ps1"
    } else {
        Write-Err "Module import failed"
        exit 1
    }
} catch {
    Write-Err "Module import test failed: $($_.Exception.Message)"
    exit 1
}

Write-Ok "Installation complete!"
