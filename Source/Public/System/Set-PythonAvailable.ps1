<#
.SYNOPSIS
    Set-PythonAvailable function from PSDocling module
.DESCRIPTION
    Sets the Python availability status in the module
.NOTES
    Part of PSDocling Document Processing System
#>
function Set-PythonAvailable {
    param(
        [bool]$Available = $true
    )

    $script:DoclingSystem.PythonAvailable = $Available
}