<#
.SYNOPSIS
    Resolves PSDocling runtime paths
.DESCRIPTION
    All runtime state (queue, status, uploads, outputs, logs, pid and helper
    scripts) lives under one per-user home so Uninstall can remove it and the
    repo stays free of live data. Default home is %LOCALAPPDATA%\PSDocling;
    set PSDOCLING_HOME to relocate it (tests use this for isolation).
    Resolved on every call so child processes and tests see the same root.
.NOTES
    Part of PSDocling Document Processing System
#>
function Get-DoclingPath {
    param(
        [ValidateSet('Home', 'Data', 'Queue', 'Output', 'Uploads', 'Logs', 'Run', 'QueueFile', 'StatusFile')]
        [string]$Name = 'Home',
        [switch]$Ensure
    )

    $root = if ($env:PSDOCLING_HOME) { $env:PSDOCLING_HOME } else { Join-Path $env:LOCALAPPDATA 'PSDocling' }
    $data = Join-Path $root 'data'

    $path = switch ($Name) {
        'Home'       { $root }
        'Data'       { $data }
        'Queue'      { Join-Path $data 'queue' }
        'Output'     { Join-Path $data 'output' }
        'Uploads'    { Join-Path $data 'uploads' }
        'Logs'       { Join-Path $root 'logs' }
        'Run'        { Join-Path $root 'run' }
        'QueueFile'  { Join-Path $data 'queue.json' }
        'StatusFile' { Join-Path $data 'status.json' }
    }

    if ($Ensure -and $Name -notin @('QueueFile', 'StatusFile') -and -not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
    return $path
}
