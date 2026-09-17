#!/usr/bin/env powershell
<#
.SYNOPSIS
    Uninstalls the PSDocling PowerShell module
.DESCRIPTION
    Removes the PSDocling module from the PowerShell module directory.
.PARAMETER Scope
    Installation scope to remove from: CurrentUser, AllUsers, or Both (default: CurrentUser)
.PARAMETER Force
    Force removal without confirmation prompts
.EXAMPLE
    .\scripts\Uninstall-DoclingModule.ps1
.EXAMPLE
    .\scripts\Uninstall-DoclingModule.ps1 -Scope Both -Force
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

    if ($ScopeName -eq 'AllUsers') {
        $destBase = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
        $destBasePSCore = Join-Path $env:ProgramFiles 'PowerShell\Modules'
    } else {
        $destBase = Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules'
        $destBasePSCore = Join-Path $env:USERPROFILE 'Documents\PowerShell\Modules'
    }

    $moduleName = 'PSDocling'
    $destDir = Join-Path $destBase $moduleName
    $destDirPSCore = Join-Path $destBasePSCore $moduleName
    $removed = $false

    if (Test-Path $destDir) {
        Write-Info "Found module at: $destDir"
        if ($Force -or (Read-Host "Remove module from $destDir? (y/N)") -eq 'y') {
            try {
                Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
                Remove-Item $destDir -Recurse -Force
                Write-Ok "Removed: $destDir"
                $removed = $true
            } catch {
                Write-Err "Failed to remove $destDir`: $($_.Exception.Message)"
            }
        }
    }

    if (Test-Path $destDirPSCore) {
        Write-Info "Found module at: $destDirPSCore"
        if ($Force -or (Read-Host "Remove module from $destDirPSCore? (y/N)") -eq 'y') {
            try {
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

if (($Scope -eq 'AllUsers' -or $Scope -eq 'Both') -and -not (Test-IsAdmin)) {
    Write-Err "Removing from AllUsers scope requires administrator privileges. Run PowerShell as Administrator."
    exit 1
}

Write-Info "Uninstalling PSDocling module..."

try {
    Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
    Write-Info "Removed module from current session"
} catch { }

switch ($Scope) {
    'CurrentUser' { Remove-ModuleFromScope 'CurrentUser' }
    'AllUsers'    { Remove-ModuleFromScope 'AllUsers' }
    'Both' {
        Remove-ModuleFromScope 'CurrentUser'
        Remove-ModuleFromScope 'AllUsers'
    }
}

Write-Ok "Uninstall complete!"
