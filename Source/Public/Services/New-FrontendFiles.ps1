<#
.SYNOPSIS
    Copies the PSDocling web UI to a folder
.DESCRIPTION
    DoclingFrontend\index.html is the single source of the UI and is served
    by the API server. This command copies that folder to -Destination (for
    example to customise the UI). It used to embed a second copy of the HTML,
    which drifted from the real file and overwrote it on -GenerateFrontend.
.PARAMETER Destination
    Target folder. Defaults to .\DoclingFrontend in the current directory.
.NOTES
    Part of PSDocling Document Processing System
#>
function New-FrontendFiles {
    [CmdletBinding()]
    param(
        [string]$Destination = (Join-Path (Get-Location) 'DoclingFrontend')
    )

    $source = Get-DoclingFrontendPath
    if (-not $source) {
        Write-Warning "DoclingFrontend not found next to the module. Reinstall PSDocling."
        return $null
    }

    $sourceFull = [System.IO.Path]::GetFullPath($source).TrimEnd('\')
    $destFull = [System.IO.Path]::GetFullPath($Destination).TrimEnd('\')
    if ($sourceFull -eq $destFull) {
        Write-Verbose "Frontend already at $destFull"
        return $destFull
    }

    New-Item -ItemType Directory -Path $destFull -Force | Out-Null
    Copy-Item -Path (Join-Path $sourceFull '*') -Destination $destFull -Recurse -Force
    Write-Host "Frontend copied to: $destFull" -ForegroundColor Green
    return $destFull
}
