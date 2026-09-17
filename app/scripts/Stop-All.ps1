<#
.SYNOPSIS
    Stops all PSDocling processes for the current data folder
.PARAMETER ClearQueue
    Also clear the queue and processing history (finished output files are kept
    unless you confirm).
#>
param([switch]$ClearQueue)

$appDir = Split-Path -Parent $PSScriptRoot
$candidates = @(
    (Join-Path $PSScriptRoot 'PSDocling.psm1'),
    (Join-Path $appDir 'Build\PSDocling.psm1')
)
$modulePath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($modulePath) {
    Import-Module $modulePath -Force
} elseif (Get-Module -ListAvailable PSDocling) {
    Import-Module PSDocling -Force
} else {
    Write-Host "PSDocling module not found." -ForegroundColor Red
    exit 1
}

Stop-DoclingSystem -ClearQueue:$ClearQueue
