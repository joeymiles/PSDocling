<#
.SYNOPSIS
    Clear-PSDoclingSystem function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
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

    # Clear the queue file
    $queueFile = "$env:TEMP\docling_queue.json"
    if (Test-Path $queueFile) {
        "[]" | Set-Content $queueFile -Encoding UTF8
        Write-Host "Cleared queue file" -ForegroundColor Green
    }
    else {
        Write-Host "Queue file doesn't exist" -ForegroundColor Gray
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

    # Optional: Clear processed documents directory
    $processedDir = ".\ProcessedDocuments"
    if (Test-Path $processedDir) {
        $docCount = (Get-ChildItem $processedDir -Directory).Count
        if ($docCount -gt 0) {
            Write-Host "Found $docCount document folders in ProcessedDocuments" -ForegroundColor Yellow
            $clearDocs = Read-Host "Clear ProcessedDocuments folder too? (Y/N)"
            if ($clearDocs -eq 'Y') {
                Remove-Item "$processedDir\*" -Recurse -Force
                Write-Host "Cleared ProcessedDocuments" -ForegroundColor Green
            }
        }
    }

    # Optional: Clear temp processing directory
    $tempDir = "$env:TEMP\DoclingProcessor"
    if (Test-Path $tempDir) {
        $tempCount = (Get-ChildItem $tempDir -Directory -ErrorAction SilentlyContinue).Count
        if ($tempCount -gt 0) {
            Write-Host "Found $tempCount temp folders in DoclingProcessor" -ForegroundColor Yellow
            $clearTemp = Read-Host "Clear temp processing folders? (Y/N)"
            if ($clearTemp -eq 'Y') {
                Remove-Item "$tempDir\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Cleared temp processing folders" -ForegroundColor Green
            }
        }
    }

    Write-Host "`nSystem cleared!" -ForegroundColor Green
    Write-Host "You can now restart the system with: .\Start-All.ps1" -ForegroundColor Cyan
}
