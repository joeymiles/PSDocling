<#
.SYNOPSIS
    Get-SecureFileName function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-SecureFileName {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,
        [string[]]$AllowedExtensions = @('.pdf', '.docx', '.xlsx', '.pptx', '.md', '.html', '.xhtml', '.csv', '.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.webp', '.txt')
    )

    if (-not (Test-SecureFileName -FileName $FileName -AllowedExtensions $AllowedExtensions)) {
        throw "Invalid or potentially dangerous filename: $FileName"
    }

    return [System.IO.Path]::GetFileName($FileName)
}
