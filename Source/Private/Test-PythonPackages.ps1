<#
.SYNOPSIS
    Test-PythonPackages function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Test-PythonPackages {
    param(
        [switch]$InstallMissing
    )

    $requiredPackages = @{
        'docling' = 'Document processing library'
        'docling-core' = 'Core document types'
        'transformers' = 'HuggingFace tokenizers for chunking'
        'tiktoken' = 'OpenAI tokenizers for chunking'
    }

    $missing = @()
    foreach ($package in $requiredPackages.Keys) {
        $pipShow = & python -m pip show $package 2>&1
        if (-not ($pipShow -match "Name: $package")) {
            $missing += $package
        }
    }

    if ($missing.Count -gt 0) {
        if ($InstallMissing) {
            Write-Host "Installing required Python packages..." -ForegroundColor Yellow
            foreach ($package in $missing) {
                Write-Host "  Installing $package ($($requiredPackages[$package]))..." -ForegroundColor Yellow
                & python -m pip install $package --quiet 2>$null
                if ($LASTEXITCODE -ne 0) {
                    Write-Warning "Failed to install $package"
                }
            }
            Write-Host "All required packages installed" -ForegroundColor Green
            return $true
        } else {
            Write-Warning "Missing Python packages: $($missing -join ', ')"
            Write-Host "Run Initialize-DoclingSystem to install missing packages" -ForegroundColor Yellow
            return $false
        }
    }
    return $true
}
