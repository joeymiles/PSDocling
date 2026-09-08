<#
.SYNOPSIS
    Get-PythonStatus function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-PythonStatus {
    return $script:DoclingSystem.PythonAvailable
}
