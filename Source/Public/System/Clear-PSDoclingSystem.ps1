<#
.SYNOPSIS
    Clear-PSDoclingSystem function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
.NOTES
    Part of PSDocling Document Processing System
#>
Function Clear-PSDoclingSystem {
    # Clears all queued items and processing status from the Docling system

    param(
        [switch]$Force
    )

    Write-Host "Clearing Docling System..." -ForegroundColor Cyan

    # Confirm with user unless -Force is specified
    if (-not $Force) {
        $confirm = Read-Host "This will clear all queued and processing documents. Continue? (Y/N)"
        if ($confirm -ne 'Y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Clear the queue folder (new folder-based queue)
    $queueFolder = "$env:TEMP\DoclingQueue"
    if (Test-Path $queueFolder) {
        $queueCount = @(Get-ChildItem $queueFolder -Filter "*.queue" -ErrorAction SilentlyContinue).Count
        if ($queueCount -gt 0) {
            Remove-Item "$queueFolder\*.queue" -Force
            Write-Host "Cleared $queueCount items from queue folder" -ForegroundColor Green
        }
        else {
            Write-Host "Queue folder is already empty" -ForegroundColor Gray
        }
    }
    else {
        Write-Host "Queue folder doesn't exist" -ForegroundColor Gray
    }

    # Clear the old JSON queue file (for backwards compatibility)
    $queueFile = "$env:TEMP\docling_queue.json"
    if (Test-Path $queueFile) {
        "[]" | Set-Content $queueFile -Encoding UTF8
        Write-Host "Cleared old queue file" -ForegroundColor Green
    }

    # Clear the status file
    $statusFile = "$env:TEMP\docling_status.json"
    if (Test-Path $statusFile) {
        "{}" | Set-Content $statusFile -Encoding UTF8
        Write-Host "Cleared status file" -ForegroundColor Green
    }
    else {
        Write-Host "Status file doesn't exist" -ForegroundColor Gray
    }

    # Optional: Clear output directories (module OutputDirectory + legacy relative path)
    $processedDirs = @()
    if ($script:DoclingSystem -and $script:DoclingSystem.OutputDirectory) {
        $processedDirs += $script:DoclingSystem.OutputDirectory
    }
    $processedDirs += @(".\ProcessedDocuments", "$env:TEMP\DoclingOutput")
    $processedDirs = $processedDirs | Where-Object { $_ } | Select-Object -Unique
    foreach ($processedDir in $processedDirs) {
        if (Test-Path $processedDir) {
            $docCount = @(Get-ChildItem $processedDir -Directory -ErrorAction SilentlyContinue).Count
            if ($docCount -gt 0) {
                Write-Host "Found $docCount document folders in $processedDir" -ForegroundColor Yellow
                $doClear = [bool]$Force
                if (-not $Force) {
                    $clearDocs = Read-Host "Clear $processedDir folder too? (Y/N)"
                    $doClear = ($clearDocs -eq 'Y')
                }
                if ($doClear) {
                    Remove-Item "$processedDir\*" -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "Cleared $processedDir" -ForegroundColor Green
                }
            }
        }
    }

    # Optional: Clear temp processing directory
    $tempDir = "$env:TEMP\DoclingProcessor"
    if (Test-Path $tempDir) {
        $tempCount = @(Get-ChildItem $tempDir -Directory -ErrorAction SilentlyContinue).Count
        if ($tempCount -gt 0) {
            Write-Host "Found $tempCount temp folders in DoclingProcessor" -ForegroundColor Yellow
            $doClear = [bool]$Force
            if (-not $Force) {
                $clearTemp = Read-Host "Clear temp processing folders? (Y/N)"
                $doClear = ($clearTemp -eq 'Y')
            }
            if ($doClear) {
                Remove-Item "$tempDir\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Cleared temp processing folders" -ForegroundColor Green
            }
        }
    }

    Write-Host "`nSystem cleared!" -ForegroundColor Green
    Write-Host "You can now restart the system with: .\Start-All.ps1" -ForegroundColor Cyan
}
