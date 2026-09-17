<#
.SYNOPSIS
    Locates the DoclingFrontend folder
.DESCRIPTION
    Checks the installed module folder first, then the repo checkout (sibling
    of Build/), then the current directory. Returns $null when not found.
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-DoclingFrontendPath {
    $moduleDir = if ($script:DoclingSystem -and $script:DoclingSystem.ModulePath) {
        Split-Path $script:DoclingSystem.ModulePath -Parent
    } else {
        $PSScriptRoot
    }
    $candidates = @(
        (Join-Path $moduleDir 'DoclingFrontend'),
        (Join-Path (Split-Path $moduleDir -Parent) 'DoclingFrontend'),
        (Join-Path (Get-Location) 'DoclingFrontend')
    )
    foreach ($candidate in $candidates) {
        if (Test-Path (Join-Path $candidate 'index.html')) {
            return (Resolve-Path $candidate).Path
        }
    }
    return $null
}
