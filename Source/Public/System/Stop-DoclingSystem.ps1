<#
.SYNOPSIS
    Stops all PSDocling system processes
.DESCRIPTION
    Stops the API server, document processor, and web server processes.
    Cleans up temporary files and optionally clears the queue.
.PARAMETER ClearQueue
    Also clear the queue and status files when stopping
.EXAMPLE
    Stop-DoclingSystem
    Stops all PSDocling processes
.EXAMPLE
    Stop-DoclingSystem -ClearQueue
    Stops all processes and clears the queue
.NOTES
    Part of PSDocling Document Processing System
#>
function Stop-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$ClearQueue
    )

    Write-Host "Stopping Docling System processes..." -ForegroundColor Cyan

    # Stop PowerShell processes running Docling components
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
            Write-Verbose "Found $($doclingProcesses.Count) processes from PID file"
        } catch {
            Write-Warning "Could not read PID file: $($_.Exception.Message)"
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
        Write-Host "Found $($doclingProcesses.Count) Docling processes to stop" -ForegroundColor Yellow
        $doclingProcesses | ForEach-Object {
            try {
                Write-Verbose "Stopping process $($_.Id): $($_.ProcessName)"
                $_ | Stop-Process -Force
            } catch {
                Write-Warning "Could not stop process $($_.Id): $($_.Exception.Message)"
            }
        }

        # Remove PID file after stopping processes
        if (Test-Path $pidFile) {
            Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
        }

        Write-Host "Stopped $($doclingProcesses.Count) processes" -ForegroundColor Green
    } else {
        Write-Host "No Docling processes found running" -ForegroundColor Gray
    }

    # Clean up temp files
    $tempFiles = @(
        "$env:TEMP\docling_api.ps1",
        "$env:TEMP\docling_processor.ps1",
        "$env:TEMP\docling_output.txt",
        "$env:TEMP\docling_error.txt",
        "$env:TEMP\docling_processor_debug.txt",
        "$env:TEMP\docling_processor_errors.log"
    )

    $tempFiles | ForEach-Object {
        if (Test-Path $_) {
            $retries = 3
            for ($i = 1; $i -le $retries; $i++) {
                try {
                    Remove-Item $_ -Force -ErrorAction Stop
                    Write-Verbose "Cleaned up temp file: $_"
                    break
                } catch {
                    if ($i -eq $retries) {
                        Write-Warning "Could not remove temp file: $(Split-Path $_ -Leaf)"
                    } else {
                        Start-Sleep -Milliseconds 200
                    }
                }
            }
        }
    }

    # Optionally clear queue and status
    if ($ClearQueue) {
        Write-Host "Clearing queue and status files..." -ForegroundColor Yellow
        Clear-PSDoclingSystem -Force
    }

    Write-Host "Docling System stopped" -ForegroundColor Green
}
