#Requires -Version 5.1
<#
.SYNOPSIS
    Builds PSDocling module from source files
#>
[CmdletBinding()]
param(
    [string]$OutputPath = ".\Build",
    [switch]$Install,
    [switch]$Test
)

Write-Host "`nPSDocling Module Builder" -ForegroundColor Cyan
Write-Host "========================" -ForegroundColor Cyan

# Create output directory
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# Get version
if (Get-Command Import-PowerShellDataFile -ErrorAction SilentlyContinue) {
    $manifest = Import-PowerShellDataFile ".\PSDocling.psd1"
    $version = $manifest.ModuleVersion
}
else {
    # Fallback for older PowerShell
    $version = "3.0.0"
}
Write-Host "Building version: $version"

# Start building module
$moduleContent = @'
#Requires -Version 5.1
# PSDocling Module - Built from source files

'@

# Process initialization
$initFile = ".\Source\Private\Initialize-Module.ps1"
if (Test-Path $initFile) {
    Write-Host "Adding initialization code"
    $content = Get-Content $initFile -Raw
    $content = $content -replace '(?s)^<#.*?#>\s*', ''
    $moduleContent += $content + "`n"
}

# Process private functions
Write-Host "`nProcessing private functions:"
$privateFunctions = @()
Get-ChildItem ".\Source\Private" -Filter "*.ps1" | Where-Object { $_.Name -ne 'Initialize-Module.ps1' } | ForEach-Object {
    Write-Host "  + $($_.BaseName)"
    $content = Get-Content $_.FullName -Raw
    $content = $content -replace '(?s)^<#.*?#>\s*', ''
    $moduleContent += "`n# Private: $($_.BaseName)`n$content`n"
    $privateFunctions += $_.BaseName
}

# Process public functions
Write-Host "`nProcessing public functions:"
$publicFunctions = @()
Get-ChildItem ".\Source\Public" -Filter "*.ps1" -Recurse | ForEach-Object {
    Write-Host "  + $($_.BaseName)"
    $content = Get-Content $_.FullName -Raw
    $content = $content -replace '(?s)^<#.*?#>\s*', ''
    $moduleContent += "`n# Public: $($_.BaseName)`n$content`n"
    $publicFunctions += $_.BaseName
}

# Add export
if ($publicFunctions.Count -gt 0) {
    $exportList = ($publicFunctions | ForEach-Object { "'$_'" }) -join ', '
    $moduleContent += "`nExport-ModuleMember -Function @($exportList)`n"
}

$moduleContent += "`nWrite-Host 'PSDocling Module Loaded - Version $version' -ForegroundColor Cyan`n"

# Write module
$outputModulePath = Join-Path $OutputPath "PSDocling.psm1"
Set-Content -Path $outputModulePath -Value $moduleContent -Encoding UTF8
Write-Host "`nBuilt: $outputModulePath"
Write-Host "Size: $([Math]::Round((Get-Item $outputModulePath).Length / 1KB, 2)) KB"
Write-Host "Functions: $($publicFunctions.Count) public, $($privateFunctions.Count) private"

# Copy manifest
Copy-Item ".\PSDocling.psd1" (Join-Path $OutputPath "PSDocling.psd1") -Force

# Copy config
if (Test-Path ".\PSDocling.config.psd1") {
    Copy-Item ".\PSDocling.config.psd1" (Join-Path $OutputPath "PSDocling.config.psd1") -Force
}

# Install if requested
if ($Install) {
    Write-Host "`nInstalling module..."

    $installPath = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\PSDocling"

    # Remove existing
    if (Test-Path $installPath) {
        Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
        Remove-Item $installPath -Recurse -Force
    }

    # Create and copy
    New-Item -ItemType Directory -Path $installPath -Force | Out-Null
    Copy-Item "$OutputPath\*" $installPath -Force

    Write-Host "Installed to: $installPath"
}

# Test if requested
if ($Test) {
    Write-Host "`nTesting module..."

    Import-Module "$installPath\PSDocling.psm1" -Force

    $functions = @('Initialize-DoclingSystem', 'Start-DoclingSystem', 'Get-DoclingSystemStatus')
    foreach ($func in $functions) {
        if (Get-Command $func -ErrorAction SilentlyContinue) {
            Write-Host "  + $func exists" -ForegroundColor Green
        }
        else {
            Write-Host "  - $func missing" -ForegroundColor Red
        }
    }
}

Write-Host "`nBuild complete!" -ForegroundColor Green