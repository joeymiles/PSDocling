<#
.SYNOPSIS
    Starts the PSDocling API server, document processor and UI window
.DESCRIPTION
    Starts two hidden PowerShell processes (API server, which also serves the
    web UI, and the document processor), waits for the API to answer, then
    opens the UI:
      -UseWebView   native pywebview window (falls back to an Edge/Chrome
                    --app window when pywebview is not available)
      -OpenBrowser  a normal browser tab (developer use)
    Process ids are written to <home>\run\docling_pids.json for Stop-DoclingSystem.
.PARAMETER Port
    Port for the API and UI (loopback only). Defaults to the module setting (8080).
.NOTES
    Part of PSDocling Document Processing System
#>
function Start-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$OpenBrowser,
        [switch]$UseWebView,
        [int]$Port = 0
    )

    if ($Port -gt 0) {
        $script:DoclingSystem.APIPort = $Port
        $script:DoclingSystem.WebPort = $Port
    }
    $apiPort = $script:DoclingSystem.APIPort
    $uiUrl = "http://localhost:$apiPort"
    $runDir = Get-DoclingPath Run -Ensure

    Write-Host "Starting Docling System..." -ForegroundColor Cyan
    Write-DoclingLog -Component launcher -Message "Starting on $uiUrl (home: $(Get-DoclingPath Home))"

    $pythonAvailable = if ($script:DoclingSystem.PythonAvailable) { '$true' } else { '$false' }
    $modulePath = $script:DoclingSystem.ModulePath -replace "'", "''"

    # Child scripts log their own crash, since their consoles are hidden.
    $childTemplate = @'
$ErrorActionPreference = 'Continue'
try {
    Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
    Import-Module '__MODULE__' -Force -ErrorAction Stop *> $null
    Set-PythonAvailable -Available __PYTHON__
    __COMMAND__
} catch {
    Write-DoclingLog -Component __COMPONENT__ -Level ERROR -Message ("__COMPONENT__ crashed: " + $_.Exception.Message)
    throw
}
'@
    $childTemplate = $childTemplate.Replace('__MODULE__', $modulePath).Replace('__PYTHON__', $pythonAvailable)
    $childArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File')

    $apiPath = Join-Path $runDir 'docling_api.ps1'
    # With a window, the API stops everything once the UI has gone quiet
    # (window closed without Quit). Headless/CLI use keeps running.
    $idle = if ($UseWebView) { 180 } else { 0 }
    $childTemplate.Replace('__COMMAND__', "Start-APIServer -Port $apiPort -IdleShutdownSeconds $idle").Replace('__COMPONENT__', 'api') |
        Set-Content $apiPath -Encoding UTF8
    $apiProcess = Start-Process powershell -ArgumentList ($childArgs + "`"$apiPath`"") -PassThru -WindowStyle Hidden
    Write-Host "API server started on port $apiPort" -ForegroundColor Green

    $procPath = Join-Path $runDir 'docling_processor.ps1'
    $childTemplate.Replace('__COMMAND__', 'Start-DocumentProcessor').Replace('__COMPONENT__', 'processor') |
        Set-Content $procPath -Encoding UTF8
    $procProcess = Start-Process powershell -ArgumentList ($childArgs + "`"$procPath`"") -PassThru -WindowStyle Hidden
    Write-Host "Document processor started" -ForegroundColor Green

    # Wait for the API before opening any window
    $ready = $false
    for ($i = 0; $i -lt 60; $i++) {
        if ($apiProcess.HasExited) { break }
        try {
            Invoke-RestMethod "$uiUrl/api/health" -TimeoutSec 2 -ErrorAction Stop | Out-Null
            $ready = $true
            break
        } catch {
            Start-Sleep -Milliseconds 500
        }
    }
    if (-not $ready) {
        $msg = "API did not start on $uiUrl (port in use or startup error). See $(Join-Path (Get-DoclingPath Logs) 'api.log')"
        Write-DoclingLog -Component launcher -Level ERROR -Message $msg
        Set-Content -Path (Join-Path (Get-DoclingPath Logs -Ensure) 'last-launch-error.txt') -Value "$(Get-Date -Format s) $msg" -Encoding UTF8
        Write-Warning $msg
    }

    $pyProcess = $null
    $windowProcess = $null
    if ($ready -and $UseWebView) {
        $pyProcess = Start-DoclingWindow -Url $uiUrl -ApiPort $apiPort
        if (-not $pyProcess) {
            $windowProcess = Start-DoclingAppWindow -Url $uiUrl
        }
    } elseif ($ready -and $OpenBrowser) {
        Start-Process $uiUrl
        Write-Host "Frontend opened in browser: $uiUrl" -ForegroundColor Green
    }

    Write-Host "System running at $uiUrl" -ForegroundColor Green

    $pids = @{
        API       = $apiProcess.Id
        Processor = $procProcess.Id
        PyWebView = if ($pyProcess) { $pyProcess.Id } else { $null }
        Port      = $apiPort
        Timestamp = Get-Date
    }
    $pids | ConvertTo-Json | Set-Content (Join-Path $runDir 'docling_pids.json') -Encoding UTF8

    return @{
        API       = $apiProcess
        Processor = $procProcess
        PyWebView = $pyProcess
        AppWindow = $windowProcess
        Ready     = $ready
        Url       = $uiUrl
    }
}
