<#
.SYNOPSIS
    Desktop shortcut and small app settings helpers
.DESCRIPTION
    Get-DoclingLauncherPath    Start-PSDocling.ps1 next to the installed module,
                               or app\scripts\ when running from a repo build.
    Get-DoclingShortcutPath    %USERPROFILE%\Desktop\PSDocling.lnk (known folder).
    New-DoclingShortcut        Creates the shortcut: hidden PowerShell, real icon.
    Get/Set-DoclingSetting     data\settings.json (e.g. ShortcutOffered).
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-DoclingLauncherPath {
    $moduleDir = Split-Path $script:DoclingSystem.ModulePath -Parent
    $candidates = @(
        (Join-Path $moduleDir 'Start-PSDocling.ps1'),
        (Join-Path (Split-Path $moduleDir -Parent) 'scripts\Start-PSDocling.ps1')
    )
    return $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Get-DoclingIconPath {
    $frontend = Get-DoclingFrontendPath
    if (-not $frontend) { return $null }
    $icon = Join-Path $frontend 'psdocling.ico'
    if (Test-Path $icon) { return $icon }
    return $null
}

function Get-DoclingShortcutPath {
    # PSDOCLING_DESKTOP lets tests avoid touching the real desktop
    $desktop = if ($env:PSDOCLING_DESKTOP) { $env:PSDOCLING_DESKTOP } else { [Environment]::GetFolderPath('Desktop') }
    return Join-Path $desktop 'PSDocling.lnk'
}

function New-DoclingShortcut {
    $launcher = Get-DoclingLauncherPath
    if (-not $launcher) { throw "Start-PSDocling.ps1 not found next to the module" }
    $shortcutPath = Get-DoclingShortcutPath
    $shell = New-Object -ComObject WScript.Shell
    $lnk = $shell.CreateShortcut($shortcutPath)
    $lnk.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $lnk.Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$launcher`""
    $lnk.WorkingDirectory = Split-Path $launcher -Parent
    $lnk.WindowStyle = 7
    $lnk.Description = 'PSDocling document processor'
    $icon = Get-DoclingIconPath
    if ($icon) { $lnk.IconLocation = "$icon,0" }
    $lnk.Save()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    Write-DoclingLog -Component launcher -Message "Desktop shortcut created: $shortcutPath"
    return $shortcutPath
}

function Get-DoclingSetting {
    param([Parameter(Mandatory)][string]$Name)
    $file = Join-Path (Get-DoclingPath Data) 'settings.json'
    if (-not (Test-Path $file)) { return $null }
    try {
        $settings = Get-Content $file -Raw | ConvertFrom-Json
        return $settings.$Name
    } catch {
        return $null
    }
}

function Set-DoclingSetting {
    param([Parameter(Mandatory)][string]$Name, $Value)
    $file = Join-Path (Get-DoclingPath Data -Ensure) 'settings.json'
    $settings = @{}
    if (Test-Path $file) {
        try {
            (Get-Content $file -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $settings[$_.Name] = $_.Value }
        } catch { }
    }
    $settings[$Name] = $Value
    $settings | ConvertTo-Json | Set-Content $file -Encoding UTF8
}
