<#
.SYNOPSIS
    Start-DoclingSystem function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
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

    # Start web server
    $webPath = ".\DoclingFrontend\Start-WebServer.ps1"
    if (Test-Path $webPath) {
        $webProcess = Start-Process powershell -ArgumentList "-File", $webPath, "-Port", $script:DoclingSystem.WebPort -PassThru -WindowStyle Hidden
        Write-Host "Web server started on port $($script:DoclingSystem.WebPort)" -ForegroundColor Green

        if ($UseWebView) {
            Start-Sleep 2
            # Try multiple locations for the PyWebView script
            $pyWebViewScript = $null
            $searchPaths = @(
                ".\Launch-PyWebView.py",                          # Current directory
                (Join-Path $PSScriptRoot "..\..\..\Launch-PyWebView.py"),  # From Build folder
                (Join-Path (Split-Path $PSScriptRoot -Parent) "..\..\..\Launch-PyWebView.py")  # From nested source
            )

            foreach ($path in $searchPaths) {
                $resolvedPath = Resolve-Path $path -ErrorAction SilentlyContinue
                if ($resolvedPath -and (Test-Path $resolvedPath)) {
                    $pyWebViewScript = $resolvedPath.Path
                    break
                }
            }

            if ($pyWebViewScript) {
                $pyProcess = Start-Process python -ArgumentList $pyWebViewScript, $script:DoclingSystem.APIPort, $script:DoclingSystem.WebPort -PassThru
                Write-Host "PyWebView window launched" -ForegroundColor Green
            } else {
                Write-Warning "PyWebView script not found. Install pywebview with: pip install pywebview requests"
                Write-Host "Falling back to browser mode" -ForegroundColor Yellow
                Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
            }
        } elseif ($OpenBrowser) {
            Start-Sleep 2
            Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
        }
    }

    Write-Host "System running! Frontend: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green

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
