<#
.SYNOPSIS
    Opens the PSDocling UI in its own window
.DESCRIPTION
    Start-DoclingWindow launches Launch-PyWebView.py with pythonw.exe (no
    console). It returns $null when pywebview cannot be used so the caller
    can fall back to Start-DoclingAppWindow, which opens Edge or Chrome in
    --app mode with a private profile under <home>\run.
.NOTES
    Part of PSDocling Document Processing System
#>
function Start-DoclingWindow {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][int]$ApiPort
    )

    $moduleDir = Split-Path $script:DoclingSystem.ModulePath -Parent
    $script = @(
        (Join-Path $moduleDir 'Launch-PyWebView.py'),
        (Join-Path (Split-Path $moduleDir -Parent) 'scripts\Launch-PyWebView.py'),
        (Join-Path (Get-Location) 'scripts\Launch-PyWebView.py')
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    $pythonw = (Get-Command pythonw -ErrorAction SilentlyContinue).Source
    if (-not $script -or -not $pythonw) {
        Write-DoclingLog -Component launcher -Level WARN -Message "pywebview launcher unavailable (script: $script, pythonw: $pythonw)"
        return $null
    }

    $python = Join-Path (Split-Path $pythonw -Parent) 'python.exe'
    if (Test-Path $python) {
        & $python -c "import webview" 2>$null
        if ($LASTEXITCODE -ne 0) {
            Write-DoclingLog -Component launcher -Level WARN -Message "pywebview not installed for $python"
            return $null
        }
    }

    $runDir = Get-DoclingPath Run -Ensure
    $icon = Join-Path (Get-DoclingFrontendPath) 'psdocling.ico'
    $proc = Start-Process $pythonw -ArgumentList "`"$script`"", $ApiPort, "`"$runDir`"", "`"$icon`"" -PassThru
    Write-DoclingLog -Component launcher -Message "pywebview window started (pid $($proc.Id))"
    Write-Host "PSDocling window launched" -ForegroundColor Green
    return $proc
}

function Start-DoclingAppWindow {
    param([Parameter(Mandatory)][string]$Url)

    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
        (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
    )
    $browser = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $browser) {
        Write-DoclingLog -Component launcher -Level WARN -Message "No Edge/Chrome found; opening default browser"
        Start-Process $Url
        return $null
    }

    $profileDir = Join-Path (Get-DoclingPath Run -Ensure) 'browser-profile'
    $proc = Start-Process $browser -ArgumentList "--app=$Url", "--user-data-dir=`"$profileDir`"", '--no-first-run', '--window-size=1400,900' -PassThru
    Write-DoclingLog -Component launcher -Message "App window started with $browser"
    Write-Host "PSDocling app window launched" -ForegroundColor Green
    return $proc
}
