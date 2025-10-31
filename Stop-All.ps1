# Stop-All.ps1 - Stops all PSDocling system processes

# Import module to ensure functions are available
# Try to use built module first, fall back to source
$buildModulePath = Join-Path $PSScriptRoot 'Build\PSDocling.psm1'
$sourceModulePath = Join-Path $PSScriptRoot 'PSDocling.psm1'

if (Test-Path $buildModulePath) {
    $modulePath = $buildModulePath
} elseif (Test-Path $sourceModulePath) {
    $modulePath = $sourceModulePath
} else {
    Write-Host "PSDocling module not found in Build or root folder" -ForegroundColor Red
    exit 1
}

Import-Module $modulePath -Force

# Stop the system using the module function
Stop-DoclingSystem -ClearQueue
