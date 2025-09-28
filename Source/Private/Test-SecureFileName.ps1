<#
.SYNOPSIS
    Test-SecureFileName function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
function Test-SecureFileName {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,
        [string[]]$AllowedExtensions = @('.pdf', '.docx', '.xlsx', '.pptx', '.md', '.html', '.xhtml', '.csv', '.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.webp', '.txt')
    )

    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return $false
    }

    # Extract just the base filename to prevent directory traversal
    $safeFileName = [System.IO.Path]::GetFileName($FileName)

    # Additional security checks
    if ($safeFileName -ne $FileName) {
        # Original contained path separators - potential traversal attempt
        return $false
    }

    # Check for invalid characters (beyond what GetFileName handles)
    if ($safeFileName -match '[<>:"|?*]') {
        return $false
    }

    # Check length limits (NTFS limit is 255, but we'll be more conservative)
    if ($safeFileName.Length -gt 200) {
        return $false
    }

    # Check extension if provided
    if ($AllowedExtensions.Count -gt 0) {
        $extension = [System.IO.Path]::GetExtension($safeFileName).ToLower()
        if ($extension -notin $AllowedExtensions) {
            return $false
        }
    }

    # Check for reserved names (Windows)
    $reservedNames = @('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($safeFileName).ToUpper()
    if ($nameWithoutExt -in $reservedNames) {
        return $false
    }

    return $true
}
