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
        [switch]$OpenBrowser
    )

    Write-Host "Starting Docling System..." -ForegroundColor Cyan

    # Start API server
    # Pass Python availability status to subprocess
    $pythonAvailable = if ($script:DoclingSystem.PythonAvailable) { '$true' } else { '$false' }
    $apiScript = @"
Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
Import-Module '$PSCommandPath' -Force
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
Import-Module '$PSCommandPath' -Force
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

        if ($OpenBrowser) {
            Start-Sleep 2
            Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
        }
    }

    Write-Host "System running! Frontend: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green

    return @{
        API       = $apiProcess
        Processor = $procProcess
        Web       = $webProcess
    }
}
