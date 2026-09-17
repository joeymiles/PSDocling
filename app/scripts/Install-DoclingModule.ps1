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
.PARAMETER DesktopShortcut
    Create the desktop shortcut without asking
.PARAMETER NoDesktopShortcut
    Do not create or offer the desktop shortcut
.EXAMPLE
    .\app\scripts\Install-DoclingModule.ps1
.EXAMPLE
    .\app\scripts\Install-DoclingModule.ps1 -Force -DesktopShortcut
#>
param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [switch]$Force,
    [switch]$SkipBuild,
    [switch]$DesktopShortcut,
    [switch]$NoDesktopShortcut,
    # Tests only: install into this folder instead of the PowerShell module paths
    [string]$Destination
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
if ($Destination) { $destDirs = @((Join-Path $Destination $moduleName)) }

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

    # Window launcher, desktop launcher, uninstaller and stop helper run from the module folder
    foreach ($script in @('Launch-PyWebView.py', 'Start-PSDocling.ps1', 'Uninstall-PSDocling.ps1', 'Stop-All.ps1')) {
        $scriptSrc = Join-Path $PSScriptRoot $script
        if (Test-Path $scriptSrc) {
            Copy-Item $scriptSrc (Join-Path $DestDir $script) -Force
            Write-Info "Copied: $script -> $DestDir"
        }
    }

    foreach ($file in @('README.md', 'LICENSE', 'requirements-webview.txt')) {
        # README/LICENSE sit at the repo root, one level above app/
        $sourcePath = @((Join-Path $RepoRoot $file), (Join-Path (Split-Path $RepoRoot -Parent) $file)) |
            Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($sourcePath) {
            Copy-Item $sourcePath (Join-Path $DestDir $file) -Force
            Write-Info "Copied: $file -> $DestDir"
        }
    }
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
    } else {
        Write-Err "Module import failed"
        exit 1
    }
} catch {
    Write-Err "Module import test failed: $($_.Exception.Message)"
    exit 1
}

# Offer the desktop shortcut (the app also offers it once on first run)
$module = Get-Module PSDocling
$wantShortcut = $false
$decided = $true
if ($DesktopShortcut) {
    $wantShortcut = $true
} elseif ($NoDesktopShortcut) {
    $wantShortcut = $false
} elseif ([Environment]::UserInteractive) {
    $wantShortcut = (Read-Host "Add a PSDocling shortcut to your desktop? (Y/n)") -notmatch '^(n|no)$'
} else {
    $decided = $false   # unattended: leave the offer to the app's first run
}
if ($module -and $decided) {
    try {
        & $module {
            param($Create)
            Set-DoclingSetting -Name 'ShortcutOffered' -Value $true
            if ($Create) { New-DoclingShortcut }
        } $wantShortcut | ForEach-Object { Write-Ok "Desktop shortcut: $_" }
    } catch {
        Write-Warn "Could not create the desktop shortcut: $($_.Exception.Message)"
    }
}

Write-Ok "Installation complete!"
Write-Info ""
Write-Info "Start PSDocling from the desktop shortcut, or run:"
Write-Info "  & '$(Join-Path $primaryDest 'Start-PSDocling.ps1')'"
Write-Info "Uninstall (warns, then deletes data) from Settings in the app, or run:"
Write-Info "  & '$(Join-Path $primaryDest 'Uninstall-PSDocling.ps1')'"
