# Stop-All.ps1 - Stops all PSDocling system processes

$RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $RepoRoot 'Source')) -and -not (Test-Path (Join-Path $RepoRoot 'Build'))) {
    $RepoRoot = $PSScriptRoot
}

$buildModulePath = Join-Path $RepoRoot 'Build\PSDocling.psm1'
$installed = Get-Module -ListAvailable PSDocling -ErrorAction SilentlyContinue | Select-Object -First 1

if ($installed) {
    Import-Module PSDocling -Force
} elseif (Test-Path $buildModulePath) {
    Import-Module $buildModulePath -Force
} else {
    Write-Host "PSDocling module not found. Install with: .\scripts\Install-DoclingModule.ps1" -ForegroundColor Red
    exit 1
}

Stop-DoclingSystem -ClearQueue
