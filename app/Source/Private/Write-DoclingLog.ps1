<#
.SYNOPSIS
    Appends a line to a capped PSDocling log file
.DESCRIPTION
    Writes "timestamp [LEVEL] message" to <home>\logs\<Component>.log.
    When the file passes MaxBytes it is rotated to <Component>.log.1 (one
    backup kept), so logs never grow without bound. Logging must never break
    the caller, so all failures are swallowed.
.NOTES
    Part of PSDocling Document Processing System
#>
function Write-DoclingLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('api', 'processor', 'web', 'launcher', 'uninstall')]
        [string]$Component = 'api',
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',
        [long]$MaxBytes = 2MB
    )

    try {
        $logDir = Get-DoclingPath Logs -Ensure
        $logFile = Join-Path $logDir "$Component.log"
        if ((Test-Path $logFile) -and (Get-Item $logFile).Length -gt $MaxBytes) {
            Move-Item $logFile "$logFile.1" -Force
        }
        $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
        Add-Content -Path $logFile -Value $line -Encoding UTF8
    } catch {
        # Logging is best effort
    }
}
