#!/usr/bin/env powershell

<#
.SYNOPSIS
    Installs the PSDocling PowerShell module
.DESCRIPTION
    This script installs the PSDocling module either for the current user or all users.
    It copies the module files to the appropriate PowerShell module directory.
.PARAMETER Scope
    Installation scope: CurrentUser or AllUsers (default: CurrentUser)
.PARAMETER Force
    Force installation even if module already exists
.EXAMPLE
    .\Install-DoclingModule.ps1
    Installs for current user
.EXAMPLE
    .\Install-DoclingModule.ps1 -Scope AllUsers
    Installs for all users (requires admin)
.EXAMPLE
    .\Install-DoclingModule.ps1 -Force
    Force reinstall for current user
#>

param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [switch]$Force
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

# Check admin rights for AllUsers scope
if ($Scope -eq 'AllUsers' -and -not (Test-IsAdmin)) {
    Write-Err "Installing for AllUsers requires administrator privileges. Run PowerShell as Administrator or use -Scope CurrentUser"
    exit 1
}

# Get source directory
$sourceDir = $PSScriptRoot
$moduleName = 'PSDocling'

# Check if Build directory exists with built module
$buildDir = Join-Path $sourceDir 'Build'
$builtModule = Join-Path $buildDir 'PSDocling.psm1'
$usingBuildDir = Test-Path $builtModule

# Verify required files exist (check Build folder first, then root)
$requiredFiles = @('PSDocling.psm1', 'PSDocling.psd1')
foreach ($file in $requiredFiles) {
    $buildFilePath = Join-Path $buildDir $file
    $rootFilePath = Join-Path $sourceDir $file

    if ($usingBuildDir) {
        if (-not (Test-Path $buildFilePath)) {
            Write-Err "Required file not found in Build folder: $file"
            Write-Info "Run .\Build-PSDoclingModule.ps1 first to create the built module"
            exit 1
        }
    } else {
        if (-not (Test-Path $rootFilePath)) {
            Write-Err "Required file not found: $file"
            Write-Info "Run .\Build-PSDoclingModule.ps1 to build the module first"
            exit 1
        }
    }
}

# Get destination directory
if ($Scope -eq 'AllUsers') {
    $destBase = $env:ProgramFiles + '\WindowsPowerShell\Modules'
} else {
    $destBase = $env:USERPROFILE + '\Documents\WindowsPowerShell\Modules'
}

# For PowerShell Core, use different path
if ($PSVersionTable.PSEdition -eq 'Core') {
    if ($Scope -eq 'AllUsers') {
        $destBase = $env:ProgramFiles + '\PowerShell\Modules'
    } else {
        $destBase = $env:USERPROFILE + '\Documents\PowerShell\Modules'
    }
}

$destDir = Join-Path $destBase $moduleName

Write-Info "Installing PSDocling module..."
Write-Info "Source: $sourceDir"
Write-Info "Destination: $destDir"
Write-Info "Scope: $Scope"

# Check if module already exists
if (Test-Path $destDir) {
    if ($Force) {
        Write-Warn "Module directory exists, removing due to -Force flag..."
        Remove-Item $destDir -Recurse -Force
    } else {
        Write-Err "Module already installed at $destDir"
        Write-Info "Use -Force to overwrite or uninstall first with: Remove-Module PSDocling; Remove-Item '$destDir' -Recurse"
        exit 1
    }
}

# Create destination directory
Write-Info "Creating module directory..."
New-Item -Path $destDir -ItemType Directory -Force | Out-Null

# Use Build directory if available
if ($usingBuildDir) {
    Write-Info "Using built module from Build directory..."

    # Copy files from Build directory
    $filesToCopy = @(
        'PSDocling.psm1',
        'PSDocling.psd1',
        'PSDocling.config.psd1'
    )

    foreach ($file in $filesToCopy) {
        $sourcePath = Join-Path $buildDir $file
        if (Test-Path $sourcePath) {
            $destPath = Join-Path $destDir $file
            Copy-Item $sourcePath $destPath -Force
            Write-Info "Copied: $file (from Build)"
        }
    }

    # Copy additional files from root
    $additionalFiles = @(
        'Start-All.ps1',
        'Stop-All.ps1',
        'HowTo.ps1',
        'README.md',
        'LICENSE'
    )

    foreach ($file in $additionalFiles) {
        $sourcePath = Join-Path $sourceDir $file
        if (Test-Path $sourcePath) {
            $destPath = Join-Path $destDir $file
            Copy-Item $sourcePath $destPath -Force
            Write-Info "Copied: $file"
        }
    }
} else {
    Write-Warn "No built module found in Build directory. Run .\Build-PSDoclingModule.ps1 first!"

    # Fall back to copying source files
    Write-Info "Falling back to source structure..."
    $filesToCopy = @(
        'PSDocling.psm1',
        'PSDocling.psd1',
        'PSDocling.config.psd1',
        'Start-All.ps1',
        'Stop-All.ps1',
        'HowTo.ps1',
        'CLAUDE.md'
    )

    foreach ($file in $filesToCopy) {
        $sourcePath = Join-Path $sourceDir $file
        if (Test-Path $sourcePath) {
            $destPath = Join-Path $destDir $file
            Copy-Item $sourcePath $destPath -Force
            Write-Info "Copied: $file"
        } else {
            Write-Warn "Optional file not found, skipping: $file"
        }
    }

    # Copy Source directory with all function files
    $sourceSubDir = Join-Path $sourceDir 'Source'
    if (Test-Path $sourceSubDir) {
        $destSourceDir = Join-Path $destDir 'Source'
        Write-Info "Copying Source directory with function files..."
        Copy-Item $sourceSubDir $destSourceDir -Recurse -Force
        $functionCount = (Get-ChildItem -Path $sourceSubDir -Filter "*.ps1" -Recurse).Count
        Write-Info "Copied: Source/ ($functionCount function files)"
    } else {
        Write-Warn "Source directory not found - module may not work properly"
    }
}

# Test module import
Write-Info "Testing module import..."
try {
    Import-Module $destDir -Force
    $moduleInfo = Get-Module PSDocling
    if ($moduleInfo) {
        Write-Ok "Module installed successfully!"
        Write-Ok "Version: $($moduleInfo.Version)"
        Write-Ok "Functions exported: $($moduleInfo.ExportedFunctions.Count)"

        Write-Info ""
        Write-Info "Usage:"
        Write-Info "  Import-Module PSDocling"
        Write-Info "  Initialize-DoclingSystem"
        Write-Info "  Start-DoclingSystem"
        Write-Info ""
        Write-Info "Or use the convenience scripts:"
        Write-Info "  .\Start-All.ps1"
        Write-Info "  .\Stop-All.ps1"
    } else {
        Write-Err "Module import failed"
        exit 1
    }
} catch {
    Write-Err "Module import test failed: $($_.Exception.Message)"
    exit 1
}

Write-Ok "Installation complete!"