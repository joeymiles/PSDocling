<#
.SYNOPSIS
    Start-DoclingSystem function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
.NOTES
    Part of PSDocling Document Processing System
#>
function Start-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$OpenBrowser,
        [switch]$UseWebView
    )

    Write-Host "Starting Docling System..." -ForegroundColor Cyan

    # Start API server
    # Pass Python availability status to subprocess
    $pythonAvailable = if ($script:DoclingSystem.PythonAvailable) { '$true' } else { '$false' }
    $modulePath = $script:DoclingSystem.ModulePath
    $apiScript = @"
Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
Import-Module '$modulePath' -Force
Set-PythonAvailable -Available $pythonAvailable
Start-APIServer -Port $($script:DoclingSystem.APIPort)
"@
    $apiPath = Join-Path $env:TEMP "docling_api.ps1"
    $apiScript | Set-Content $apiPath -Encoding UTF8

    $apiProcess = Start-Process powershell -ArgumentList "-File", $apiPath -PassThru -WindowStyle Hidden
    Write-Host "API server started on port $($script:DoclingSystem.APIPort)" -ForegroundColor Green

    # Start processor
    # Pass Python availability status to subprocess
    $procScript = @"
Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
Import-Module '$modulePath' -Force
Set-PythonAvailable -Available $pythonAvailable
Start-DocumentProcessor
"@
    $procPath = Join-Path $env:TEMP "docling_processor.ps1"
    $procScript | Set-Content $procPath -Encoding UTF8

    $procProcess = Start-Process powershell -ArgumentList "-File", $procPath -PassThru -WindowStyle Hidden
    Write-Host "Document processor started" -ForegroundColor Green

    # Resolve frontend: installed module dir, repo (sibling of Build/), then cwd
    $frontendCandidates = @(
        (Join-Path $PSScriptRoot 'DoclingFrontend'),
        (Join-Path (Split-Path $PSScriptRoot -Parent) 'DoclingFrontend'),
        (Join-Path (Get-Location) 'DoclingFrontend')
    )
    $frontendDir = $null
    foreach ($candidate in $frontendCandidates) {
        if (Test-Path $candidate) {
            $frontendDir = $candidate
            break
        }
    }
    $webPath = if ($frontendDir) { Join-Path $frontendDir 'Start-WebServer.ps1' } else { $null }

    if ($webPath -and (Test-Path $webPath)) {
        $webProcess = Start-Process powershell -ArgumentList "-File", $webPath, "-Port", $script:DoclingSystem.WebPort -PassThru -WindowStyle Hidden
        Write-Host "Web server started on port $($script:DoclingSystem.WebPort)" -ForegroundColor Green

        if ($UseWebView) {
            Start-Sleep 2
            $pyWebViewScript = $null
            $searchPaths = @(
                (Join-Path $PSScriptRoot 'Launch-PyWebView.py'),
                (Join-Path (Split-Path $PSScriptRoot -Parent) 'scripts\Launch-PyWebView.py'),
                (Join-Path (Get-Location) 'scripts\Launch-PyWebView.py'),
                (Join-Path (Get-Location) 'Launch-PyWebView.py')
            )

            foreach ($path in $searchPaths) {
                if ($path -and (Test-Path $path)) {
                    $pyWebViewScript = (Resolve-Path $path).Path
                    break
                }
            }

            if ($pyWebViewScript) {
                # Use pythonw.exe to launch without console window (GUI only)
                $pythonw = (Get-Command pythonw -ErrorAction SilentlyContinue).Source
                if (-not $pythonw) {
                    # Fallback to python.exe if pythonw not found
                    $pythonw = 'python'
                }
                $pyProcess = Start-Process $pythonw -ArgumentList $pyWebViewScript, $script:DoclingSystem.APIPort, $script:DoclingSystem.WebPort -PassThru
                Write-Host "PyWebView window launched" -ForegroundColor Green
            } else {
                Write-Warning "PyWebView script not found. Install pywebview with: pip install pywebview requests"
                Write-Host "Falling back to browser mode" -ForegroundColor Yellow
                Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
                Write-Host "Frontend opened in browser: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green
            }
        } elseif ($OpenBrowser) {
            Start-Sleep 2
            Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
            Write-Host "Frontend opened in browser: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green
        }
    } else {
        Write-Warning "DoclingFrontend not found. Run Initialize-DoclingSystem -GenerateFrontend or install from the repo."
    }

    Write-Host "System running!" -ForegroundColor Green

    # Store process IDs for reliable cleanup
    $pidFile = Join-Path $env:TEMP "docling_pids.json"
    $pids = @{
        API       = $apiProcess.Id
        Processor = $procProcess.Id
        Web       = if ($webProcess) { $webProcess.Id } else { $null }
        PyWebView = if ($pyProcess) { $pyProcess.Id } else { $null }
        Timestamp = Get-Date
    }
    $pids | ConvertTo-Json | Set-Content $pidFile -Encoding UTF8

    return @{
        API       = $apiProcess
        Processor = $procProcess
        Web       = $webProcess
        PyWebView = if ($pyProcess) { $pyProcess } else { $null }
    }
}
