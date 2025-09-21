#!/usr/bin/env powershell

<#
.SYNOPSIS
    Uninstalls the PSDocling PowerShell module
.DESCRIPTION
    This script removes the PSDocling module from the PowerShell module directory.
    It can remove from CurrentUser or AllUsers scope.
.PARAMETER Scope
    Installation scope to remove from: CurrentUser, AllUsers, or Both (default: CurrentUser)
.PARAMETER Force
    Force removal without confirmation prompts
.EXAMPLE
    .\Uninstall-DoclingModule.ps1
    Removes from current user scope
.EXAMPLE
    .\Uninstall-DoclingModule.ps1 -Scope AllUsers
    Removes from all users scope (requires admin)
.EXAMPLE
    .\Uninstall-DoclingModule.ps1 -Scope Both -Force
    Removes from both scopes without prompts
#>

param(
    [ValidateSet('CurrentUser', 'AllUsers', 'Both')]
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

function Remove-ModuleFromScope {
    param([string]$ScopeName)

    Write-Info "Checking $ScopeName scope..."

    # Get destination directory
    if ($ScopeName -eq 'AllUsers') {
        $destBase = $env:ProgramFiles + '\WindowsPowerShell\Modules'
        # Check PowerShell Core path too
        $destBasePSCore = $env:ProgramFiles + '\PowerShell\Modules'
    } else {
        $destBase = $env:USERPROFILE + '\Documents\WindowsPowerShell\Modules'
        # Check PowerShell Core path too
        $destBasePSCore = $env:USERPROFILE + '\Documents\PowerShell\Modules'
    }

    $moduleName = 'PSDocling'
    $destDir = Join-Path $destBase $moduleName
    $destDirPSCore = Join-Path $destBasePSCore $moduleName

    $removed = $false

    # Remove from Windows PowerShell path
    if (Test-Path $destDir) {
        Write-Info "Found module at: $destDir"
        if ($Force -or (Read-Host "Remove module from $destDir? (y/N)") -eq 'y') {
            try {
                # Try to remove module from memory first
                Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
                Remove-Item $destDir -Recurse -Force
                Write-Ok "Removed: $destDir"
                $removed = $true
            } catch {
                Write-Err "Failed to remove $destDir`: $($_.Exception.Message)"
            }
        }
    }

    # Remove from PowerShell Core path
    if (Test-Path $destDirPSCore) {
        Write-Info "Found module at: $destDirPSCore"
        if ($Force -or (Read-Host "Remove module from $destDirPSCore? (y/N)") -eq 'y') {
            try {
                # Try to remove module from memory first
                Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
                Remove-Item $destDirPSCore -Recurse -Force
                Write-Ok "Removed: $destDirPSCore"
                $removed = $true
            } catch {
                Write-Err "Failed to remove $destDirPSCore`: $($_.Exception.Message)"
            }
        }
    }

    if (-not $removed) {
        Write-Info "No PSDocling module found in $ScopeName scope"
    }
}

# Check admin rights for AllUsers scope
if (($Scope -eq 'AllUsers' -or $Scope -eq 'Both') -and -not (Test-IsAdmin)) {
    Write-Err "Removing from AllUsers scope requires administrator privileges. Run PowerShell as Administrator."
    exit 1
}

Write-Info "Uninstalling PSDocling module..."

# Remove module from memory first
try {
    Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
    Write-Info "Removed module from current session"
} catch {
    # Module not loaded, continue
}

# Remove based on scope
switch ($Scope) {
    'CurrentUser' {
        Remove-ModuleFromScope 'CurrentUser'
    }
    'AllUsers' {
        Remove-ModuleFromScope 'AllUsers'
    }
    'Both' {
        Remove-ModuleFromScope 'CurrentUser'
        Remove-ModuleFromScope 'AllUsers'
    }
}

Write-Ok "Uninstall complete!"