function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn($msg)  { Write-Host $msg -ForegroundColor Yellow }

Write-Info "Stopping Docling System processes..."

# Stop PowerShell processes running Docling components
# Use WMI to get CommandLine property (Get-Process doesn't have it in PS 5.1)
$doclingProcesses = @()

# Method 1: Check PIDs from stored file (most reliable)
$pidFile = "$env:TEMP\docling_pids.json"
if (Test-Path $pidFile) {
    try {
        $storedPids = Get-Content $pidFile | ConvertFrom-Json
        foreach ($processId in @($storedPids.API, $storedPids.Processor, $storedPids.Web)) {
            if ($processId) {
                $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
                if ($proc) {
                    $doclingProcesses += $proc
                }
            }
        }
        Write-Info "Found $($doclingProcesses.Count) processes from PID file"
    } catch {
        Write-Warn "Could not read PID file: $($_.Exception.Message)"
    }
}

# Method 2: Use WMI to search by CommandLine (slower but finds orphaned processes)
$wmiProcesses = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue
foreach ($wmiProc in $wmiProcesses) {
    if ($wmiProc.CommandLine) {
        $cmdLine = $wmiProc.CommandLine
        if ($cmdLine -like "*docling_api.ps1*" -or
            $cmdLine -like "*docling_processor.ps1*" -or
            $cmdLine -like "*Start-WebServer.ps1*") {
            $proc = Get-Process -Id $wmiProc.ProcessId -ErrorAction SilentlyContinue
            if ($proc -and $proc -notin $doclingProcesses) {
                $doclingProcesses += $proc
            }
        }
    }
}

if ($doclingProcesses) {
    Write-Info "Found $($doclingProcesses.Count) Docling processes to stop"
    $doclingProcesses | ForEach-Object {
        try {
            Write-Info "Stopping process $($_.Id): $($_.ProcessName)"
            $_ | Stop-Process -Force
        } catch {
            Write-Warn "Could not stop process $($_.Id): $($_.Exception.Message)"
        }
    }

    # Remove PID file after stopping processes
    if (Test-Path $pidFile) {
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
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