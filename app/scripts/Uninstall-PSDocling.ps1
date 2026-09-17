#Requires -Version 5.1
<#
.SYNOPSIS
    Uninstalls PSDocling and deletes its user data
.DESCRIPTION
    Always removes what the PSDocling installer owns:
      - running PSDocling processes (API, processor, window)
      - the PSDocling module folders (Windows PowerShell and PowerShell 7)
      - the PSDocling desktop shortcut (only if it points at Start-PSDocling.ps1)
    After a warning, deletes user data:
      - the data folder (%LOCALAPPDATA%\PSDocling or PSDOCLING_HOME): queue,
        history, uploads, converted output, logs and settings
      - leftovers from older versions in %TEMP% (docling_queue.json, DoclingOutput, ...)
    Optional, off by default, never removes Python itself:
      -RemoveDoclingPackages   pip uninstall docling docling-core
      -RemoveTokenizers        pip uninstall transformers tiktoken
      -RemovePyWebView         pip uninstall pywebview
    A repo checkout is not removed; delete that folder yourself.
    A log is written to %TEMP%\PSDocling-uninstall.log.
.PARAMETER Force
    Do not ask for confirmation (the in-app Uninstall has already asked).
.PARAMETER WaitForProcessId
    Wait for this process (the app's API server) to exit before removing files.
.EXAMPLE
    .\app\scripts\Uninstall-PSDocling.ps1
.EXAMPLE
    .\app\scripts\Uninstall-PSDocling.ps1 -Force -RemovePyWebView
#>
param(
    [ValidateSet('CurrentUser', 'AllUsers', 'Both')]
    [string]$Scope = 'CurrentUser',
    [switch]$Force,
    [switch]$RemoveDoclingPackages,
    [switch]$RemoveTokenizers,
    [switch]$RemovePyWebView,
    [int]$WaitForProcessId = 0,
    # Tests only: module base folders to use instead of the real Modules paths
    [string[]]$ModuleBase
)
$ErrorActionPreference = 'Continue'
$logFile = Join-Path $env:TEMP 'PSDocling-uninstall.log'
$problems = New-Object System.Collections.Generic.List[string]

function Write-Step([string]$Message, [string]$Level = 'INFO') {
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    $color = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } default { 'Cyan' } }
    Write-Host $Message -ForegroundColor $color
    if ($Level -eq 'ERROR') { $script:problems.Add($Message) }
}

function Confirm-Step([string]$Question) {
    if ($Force) { return $true }
    return (Read-Host "$Question (y/N)") -match '^(y|yes)$'
}

function Remove-PathSafely([string]$Path, [string]$What) {
    if (-not (Test-Path $Path)) { return }
    for ($i = 1; $i -le 5; $i++) {
        try {
            Remove-Item $Path -Recurse -Force -ErrorAction Stop
            Write-Step "Removed $What`: $Path"
            return
        } catch {
            if ($i -eq 5) { Write-Step "Could not remove $What at $Path`: $($_.Exception.Message)" 'ERROR' }
            else { Start-Sleep -Seconds 1 }
        }
    }
}

"" | Add-Content -Path $logFile
Write-Step "PSDocling uninstall started (scope $Scope)"

# --- Resolve what belongs to PSDocling ---
$dataHome = if ($env:PSDOCLING_HOME) { $env:PSDOCLING_HOME } else { Join-Path $env:LOCALAPPDATA 'PSDocling' }
$moduleBases = New-Object System.Collections.Generic.List[string]
if (-not $ModuleBase -and $env:PSDOCLING_MODULE_BASE) { $ModuleBase = $env:PSDOCLING_MODULE_BASE -split ';' }
if ($ModuleBase) {
    $ModuleBase | ForEach-Object { $moduleBases.Add($_) }
} elseif ($Scope -in 'CurrentUser', 'Both') {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    $moduleBases.Add((Join-Path $docs 'WindowsPowerShell\Modules'))
    $moduleBases.Add((Join-Path $docs 'PowerShell\Modules'))
}
if (-not $ModuleBase -and $Scope -in 'AllUsers', 'Both') {
    $moduleBases.Add((Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'))
    $moduleBases.Add((Join-Path $env:ProgramFiles 'PowerShell\Modules'))
}
$moduleDirs = @($moduleBases | ForEach-Object { Join-Path $_ 'PSDocling' } | Where-Object { Test-Path $_ } | Select-Object -Unique)
$desktop = if ($env:PSDOCLING_DESKTOP) { $env:PSDOCLING_DESKTOP } else { [Environment]::GetFolderPath('Desktop') }
$shortcut = Join-Path $desktop 'PSDocling.lnk'
$legacy = @('docling_queue.json', 'docling_status.json', 'docling_pids.json', 'docling_api.ps1', 'docling_processor.ps1',
            'docling_output.txt', 'docling_error.txt', 'docling_processor_debug.log', 'docling_processor_errors.log',
            'DoclingQueue', 'DoclingOutput', 'DoclingProcessor') |
    ForEach-Object { Join-Path $env:TEMP $_ } | Where-Object { Test-Path $_ }

if ($Scope -in 'AllUsers', 'Both') {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Step "Removing the AllUsers module needs an administrator PowerShell." 'ERROR'
        exit 1
    }
}

# --- Warn ---
if (-not $Force) {
    $size = 0
    if (Test-Path $dataHome) {
        $size = (Get-ChildItem $dataHome -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    }
    Write-Host ""
    Write-Host "This will uninstall PSDocling and DELETE its data:" -ForegroundColor Yellow
    Write-Host ("  Data folder: {0} ({1:N1} MB) - queue, history, uploads, converted documents, logs, settings" -f $dataHome, ($size / 1MB))
    foreach ($d in $moduleDirs) { Write-Host "  Module:      $d" }
    if (Test-Path $shortcut) { Write-Host "  Shortcut:    $shortcut" }
    Write-Host "Copy any converted documents you want to keep before continuing." -ForegroundColor Yellow
    if (-not (Confirm-Step "Uninstall PSDocling and delete this data?")) {
        Write-Step "Uninstall cancelled"
        exit 2
    }
    if (-not $RemoveDoclingPackages) { $RemoveDoclingPackages = Confirm-Step "Also remove the Docling Python packages (docling, docling-core)?" }
    if (-not $RemoveTokenizers) { $RemoveTokenizers = Confirm-Step "Also remove the tokenizer packages (transformers, tiktoken)? Other Python tools may use them" }
    if (-not $RemovePyWebView) { $RemovePyWebView = Confirm-Step "Also remove pywebview (window support)?" }
}

# --- Stop PSDocling ---
if ($WaitForProcessId -gt 0) {
    $proc = Get-Process -Id $WaitForProcessId -ErrorAction SilentlyContinue
    if ($proc) {
        Write-Step "Waiting for PSDocling (pid $WaitForProcessId) to exit"
        if (-not $proc.WaitForExit(30000)) { Stop-Process -Id $WaitForProcessId -Force -ErrorAction SilentlyContinue }
    }
}
$moduleFile = @($moduleDirs | ForEach-Object { Join-Path $_ 'PSDocling.psm1' }) +
              @((Join-Path $PSScriptRoot 'PSDocling.psm1'), (Join-Path (Split-Path $PSScriptRoot -Parent) 'Build\PSDocling.psm1')) |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if ($moduleFile) {
    try {
        Import-Module $moduleFile -Force *> $null
        Stop-DoclingSystem *> $null
        Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
        Write-Step "Stopped running PSDocling processes"
    } catch {
        Write-Step "Could not stop PSDocling cleanly: $($_.Exception.Message)" 'WARN'
    }
}
Start-Sleep -Seconds 1

# --- Installer-owned pieces ---
if (Test-Path $shortcut) {
    try {
        $lnk = (New-Object -ComObject WScript.Shell).CreateShortcut($shortcut)
        if ($lnk.Arguments -like '*Start-PSDocling.ps1*') { Remove-PathSafely $shortcut 'desktop shortcut' }
        else { Write-Step "Left $shortcut alone (it does not point at PSDocling's launcher)" 'WARN' }
    } catch {
        Write-Step "Could not inspect $shortcut`: $($_.Exception.Message)" 'WARN'
    }
}
foreach ($dir in $moduleDirs) {
    if ((Split-Path $dir -Leaf) -eq 'PSDocling' -and (Split-Path (Split-Path $dir -Parent) -Leaf) -eq 'Modules') {
        Remove-PathSafely $dir 'module'
    }
}

# --- User data ---
$homeFull = [System.IO.Path]::GetFullPath($dataHome).TrimEnd('\')
$isSafeHome = $homeFull.Length -gt 10 -and
              $homeFull -ne [System.IO.Path]::GetFullPath($env:USERPROFILE).TrimEnd('\') -and
              $homeFull -ne [System.IO.Path]::GetFullPath($env:LOCALAPPDATA).TrimEnd('\') -and
              ($env:PSDOCLING_HOME -or (Split-Path $homeFull -Leaf) -eq 'PSDocling')
if ($isSafeHome) {
    Remove-PathSafely $homeFull 'data folder'
} else {
    Write-Step "Refused to delete unexpected data folder path: $homeFull" 'ERROR'
}
foreach ($item in $legacy) { Remove-PathSafely $item 'old temp data' }

# --- Optional suggested tools ---
$packages = @()
if ($RemoveDoclingPackages) { $packages += 'docling', 'docling-core' }
if ($RemoveTokenizers) { $packages += 'transformers', 'tiktoken' }
if ($RemovePyWebView) { $packages += 'pywebview' }
if ($packages.Count -gt 0) {
    $python = (Get-Command python -ErrorAction SilentlyContinue).Source
    if ($python) {
        Write-Step "Removing Python packages: $($packages -join ', ')"
        $out = & $python -m pip uninstall -y @packages 2>&1 | Out-String
        Add-Content -Path $logFile -Value $out -Encoding UTF8
        if ($LASTEXITCODE -ne 0) { Write-Step "pip uninstall reported a problem (see log)" 'ERROR' }
    } else {
        Write-Step "Python not found; skipped removing $($packages -join ', ')" 'WARN'
    }
}

$appDir = Split-Path $PSScriptRoot -Parent
if (Test-Path (Join-Path $appDir 'Source')) {
    Write-Step "This repo checkout ($(Split-Path $appDir -Parent)) was not deleted."
}

# The in-app Uninstall runs a TEMP copy of this script; remove that copy
if ((Split-Path $PSCommandPath -Leaf) -like 'PSDocling-uninstall-*.ps1') {
    Remove-Item $PSCommandPath -Force -ErrorAction SilentlyContinue
}

if ($problems.Count -gt 0) {
    Write-Step "Uninstall finished with $($problems.Count) problem(s). Log: $logFile" 'WARN'
    exit 1
}
Write-Step "PSDocling uninstalled. Log: $logFile"
exit 0
