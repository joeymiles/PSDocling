#Requires -Version 5.1
<#
.SYNOPSIS
    Builds PSDocling module from source files into Build/
.DESCRIPTION
    Concatenates Source/ into Build/PSDocling.psm1 and copies the manifest.
    Developer tool — end users should run Install-DoclingModule.ps1 instead
    (which invokes this build automatically).
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$Install,
    [switch]$Test
)

$RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $RepoRoot 'Source'))) {
    $RepoRoot = $PSScriptRoot
}
Set-Location $RepoRoot

if (-not $OutputPath) {
    $OutputPath = Join-Path $RepoRoot 'Build'
}

Write-Host "`nPSDocling Module Builder" -ForegroundColor Cyan
Write-Host "========================" -ForegroundColor Cyan

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$manifestPath = Join-Path $RepoRoot 'PSDocling.psd1'
if (Get-Command Import-PowerShellDataFile -ErrorAction SilentlyContinue) {
    $manifest = Import-PowerShellDataFile $manifestPath
    $version = $manifest.ModuleVersion
} else {
    $version = "3.3.2"
}
Write-Host "Building version: $version"
Write-Host "Repo root: $RepoRoot"

$moduleContent = @"
#Requires -Version 5.1
# PSDocling Module - Built from source files

"@

$initFile = Join-Path $RepoRoot 'Source\Private\Initialize-Module.ps1'
if (Test-Path $initFile) {
    Write-Host "Adding initialization code"
    $content = Get-Content $initFile -Raw
    $content = $content -replace '(?s)^<#.*?#>\s*', ''
    $moduleContent += $content + "`n"
}

Write-Host "`nProcessing private functions:"
$privateFunctions = @()
Get-ChildItem (Join-Path $RepoRoot 'Source\Private') -Filter "*.ps1" |
    Where-Object { $_.Name -ne 'Initialize-Module.ps1' } |
    ForEach-Object {
        Write-Host "  + $($_.BaseName)"
        $content = Get-Content $_.FullName -Raw
        $content = $content -replace '(?s)^<#.*?#>\s*', ''
        $moduleContent += "`n# Private: $($_.BaseName)`n$content`n"
        $privateFunctions += $_.BaseName
    }

Write-Host "`nProcessing public functions:"
$publicFunctions = @()
Get-ChildItem (Join-Path $RepoRoot 'Source\Public') -Filter "*.ps1" -Recurse | ForEach-Object {
    Write-Host "  + $($_.BaseName)"
    $content = Get-Content $_.FullName -Raw
    $content = $content -replace '(?s)^<#.*?#>\s*', ''
    $moduleContent += "`n# Public: $($_.BaseName)`n$content`n"
    $publicFunctions += $_.BaseName
}

if ($publicFunctions.Count -gt 0) {
    $exportList = ($publicFunctions | ForEach-Object { "'$_'" }) -join ', '
    $moduleContent += "`nExport-ModuleMember -Function @($exportList)`n"
}

$moduleContent += "`nWrite-Host 'PSDocling Module Loaded - Version $version' -ForegroundColor Cyan`n"

$outputModulePath = Join-Path $OutputPath "PSDocling.psm1"
Set-Content -Path $outputModulePath -Value $moduleContent -Encoding UTF8
Write-Host "`nBuilt: $outputModulePath"
Write-Host "Size: $([Math]::Round((Get-Item $outputModulePath).Length / 1KB, 2)) KB"
Write-Host "Functions: $($publicFunctions.Count) public, $($privateFunctions.Count) private"

Copy-Item $manifestPath (Join-Path $OutputPath "PSDocling.psd1") -Force

$configPath = Join-Path $RepoRoot 'PSDocling.config.psd1'
if (Test-Path $configPath) {
    Copy-Item $configPath (Join-Path $OutputPath "PSDocling.config.psd1") -Force
}

if ($Install) {
    Write-Host "`nInstalling module..."
    $installPath = Join-Path $env:USERPROFILE 'Documents\WindowsPowerShell\Modules\PSDocling'
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $installPath = Join-Path $env:USERPROFILE 'Documents\PowerShell\Modules\PSDocling'
    }
    if (Test-Path $installPath) {
        Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
        Remove-Item $installPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $installPath -Force | Out-Null
    Copy-Item (Join-Path $OutputPath '*') $installPath -Force
    Write-Host "Installed to: $installPath"
}

if ($Test) {
    Write-Host "`nTesting module..."
    $testPath = if ($Install) { Join-Path $installPath 'PSDocling.psm1' } else { $outputModulePath }
    Import-Module $testPath -Force
    foreach ($func in @('Initialize-DoclingSystem', 'Start-DoclingSystem', 'Get-DoclingSystemStatus')) {
        if (Get-Command $func -ErrorAction SilentlyContinue) {
            Write-Host "  + $func exists" -ForegroundColor Green
        } else {
            Write-Host "  - $func missing" -ForegroundColor Red
        }
    }
}

Write-Host "`nBuild complete!" -ForegroundColor Green
