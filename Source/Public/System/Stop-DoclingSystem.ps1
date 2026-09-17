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
        [switch]$ClearQueue,
        # Used by the API server when it handles Quit, so it can finish its own shutdown
        [int]$ExcludeProcessId = 0
    )

    Write-Host "Stopping Docling System processes..." -ForegroundColor Cyan

    $doclingProcesses = @()
    $seenIds = @{}
    function Add-Target([int]$Id) {
        if (-not $Id -or $Id -eq $ExcludeProcessId -or $Id -eq $PID -or $seenIds.ContainsKey($Id)) { return }
        $proc = Get-Process -Id $Id -ErrorAction SilentlyContinue
        if ($proc) {
            $seenIds[$Id] = $true
            $script:doclingTargets += $proc
        }
    }
    $script:doclingTargets = @()

    # Method 1: PIDs recorded by Start-DoclingSystem for this home
    $runDir = Get-DoclingPath Run
    $pidFile = Join-Path $runDir "docling_pids.json"
    if (Test-Path $pidFile) {
        try {
            $storedPids = Get-Content $pidFile -Raw | ConvertFrom-Json
            foreach ($processId in @($storedPids.API, $storedPids.Processor, $storedPids.PyWebView)) {
                Add-Target $processId
            }
        } catch {
            Write-Warning "Could not read PID file: $($_.Exception.Message)"
        }
    }

    # Method 2: command-line match for orphans. Scoped to this home's run
    # folder so another PSDocling home (tests, second install) is untouched;
    # the pre-#44 TEMP script names are matched too so old orphans get cleaned.
    $legacyApi = Join-Path $env:TEMP 'docling_api.ps1'
    $legacyProc = Join-Path $env:TEMP 'docling_processor.ps1'
    $wmiProcesses = Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='pwsh.exe' OR Name='python.exe' OR Name='pythonw.exe' OR Name='msedge.exe' OR Name='chrome.exe'" -ErrorAction SilentlyContinue
    foreach ($wmiProc in $wmiProcesses) {
        $cmdLine = $wmiProc.CommandLine
        if (-not $cmdLine) { continue }
        $isOurs = $cmdLine.IndexOf($runDir, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                  $cmdLine.IndexOf($legacyApi, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                  $cmdLine.IndexOf($legacyProc, [StringComparison]::OrdinalIgnoreCase) -ge 0
        if ($isOurs) {
            Add-Target ([int]$wmiProc.ProcessId)
        }
    }
    $doclingProcesses = $script:doclingTargets

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
        (Join-Path $runDir "docling_api.ps1"),
        (Join-Path $runDir "docling_processor.ps1"),
        (Join-Path $runDir "docling_output.txt"),
        (Join-Path $runDir "docling_error.txt")
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
