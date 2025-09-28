function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn($msg)  { Write-Host $msg -ForegroundColor Yellow }

Write-Info "Stopping Docling System processes..."

# Stop PowerShell processes running Docling components
$processes = Get-Process powershell -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -like "*docling*" -or
    $_.CommandLine -like "*Start-APIServer*" -or
    $_.CommandLine -like "*Start-DocumentProcessor*" -or
    $_.CommandLine -like "*Start-WebServer*"
}

if ($processes) {
    $processes | ForEach-Object {
        try {
            Write-Info "Stopping process $($_.Id): $($_.ProcessName)"
            $_ | Stop-Process -Force
        } catch {
            Write-Warn "Could not stop process $($_.Id): $($_.Exception.Message)"
        }
    }
} else {
    Write-Info "No Docling processes found running."
}

# Clean up temp files
$tempFiles = @(
    "$env:TEMP\docling_api.ps1",
    "$env:TEMP\docling_processor.ps1",
    "$env:TEMP\docling_output.txt",
    "$env:TEMP\docling_error.txt"
)

$tempFiles | ForEach-Object {
    if (Test-Path $_) {
        $retries = 3
        $cleaned = $false

        for ($i = 1; $i -le $retries; $i++) {
            try {
                Remove-Item $_ -Force -ErrorAction Stop
                Write-Info "Cleaned up temp file: $_"
                $cleaned = $true
                break
            } catch {
                if ($i -eq $retries) {
                    # Only show warning on final attempt
                    Write-Warn "Could not remove temp file after $retries attempts: $(Split-Path $_ -Leaf)"
                } else {
                    # Wait briefly for file lock to release
                    Start-Sleep -Milliseconds 200
                }
            }
        }
    }
}

# Import module to ensure functions are available
# Try to use built module first, fall back to source
$buildModulePath = Join-Path $PSScriptRoot 'Build\PSDocling.psm1'
$sourceModulePath = Join-Path $PSScriptRoot 'PSDocling.psm1'

if (Test-Path $buildModulePath) {
    $modulePath = $buildModulePath
} elseif (Test-Path $sourceModulePath) {
    $modulePath = $sourceModulePath
} else {
    Write-Warn "PSDocling module not found in Build or root folder, skipping system cleanup"
    $modulePath = $null
}

if ($modulePath) {
    Import-Module $modulePath -Force
    if (Get-Command Clear-PSDoclingSystem -ErrorAction SilentlyContinue) {
        Clear-PSDoclingSystem -force $true
    }
}

Write-Ok "Docling System stopped."