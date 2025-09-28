<#
.SYNOPSIS
    Initialize-DoclingSystem function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Initialize-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$SkipPythonCheck,
        [switch]$GenerateFrontend,
        [switch]$ClearHistory
    )

    Write-Host "Initializing PS Docling System v$($script:DoclingSystem.Version)" -ForegroundColor Cyan

    # Initialize session tracking
    $script:DoclingSystem.SessionStartTime = Get-Date
    $script:DoclingSystem.SessionCompletedCount = 0

    # Create directories
    @($script:DoclingSystem.TempDirectory, $script:DoclingSystem.OutputDirectory) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            Write-Host "Created directory: $_" -ForegroundColor Green
        }
    }

    # Initialize queue and status files
    if ($ClearHistory) {
        Write-Host "Clearing processing history..." -ForegroundColor Yellow
        Set-QueueItems @()
        @{} | ConvertTo-Json | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8
        $script:DoclingSystem.HistoricalCompletedCount = 0
    } else {
        if (-not (Test-Path $script:DoclingSystem.QueueFile)) {
            Set-QueueItems @()
        }
        if (-not (Test-Path $script:DoclingSystem.StatusFile)) {
            @{} | ConvertTo-Json | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8
        }
        # Count existing completed items for historical tracking
        $existingStatus = Get-ProcessingStatus
        $script:DoclingSystem.HistoricalCompletedCount = @($existingStatus.Values | Where-Object { $_.Status -eq 'Completed' }).Count
    }

    # Check Python and install required packages
    if (-not $SkipPythonCheck) {
        try {
            $version = & python --version 2>&1
            if ($version -match "Python") {
                Write-Host "Python found: $version" -ForegroundColor Green
                $script:DoclingSystem.PythonAvailable = $true

                # Check and install all required packages
                $packagesInstalled = Test-PythonPackages -InstallMissing
                if ($packagesInstalled) {
                    Write-Host "All Python packages ready" -ForegroundColor Green
                } else {
                    Write-Warning "Some Python packages may be missing"
                }
            }
        }
        catch {
            Write-Warning "Python not found - using simulation mode"
        }
    }

    # Always generate frontend files if they don't exist
    $frontendDir = Join-Path $PSScriptRoot "DoclingFrontend"
    if ($GenerateFrontend -or -not (Test-Path $frontendDir)) {
        New-FrontendFiles
    }

    Write-Host "System initialized" -ForegroundColor Green
}
