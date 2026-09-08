#Requires -Version 5.1
<#
.SYNOPSIS
    High-confidence regression tests for PSDocling audit fixes.
.DESCRIPTION
    Runs without Docling/Python. Verifies folder-queue mutex claim behavior,
    reprocess enqueue wiring, system-status queue source, and frontend
    downloadDocument onclick string construction.
#>
$ErrorActionPreference = 'Stop'
$failed = 0
$passed = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if ($Condition) {
        Write-Host "PASS: $Message" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "FAIL: $Message" -ForegroundColor Red
        $script:failed++
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $repoRoot 'Source'))) {
    $repoRoot = (Get-Location).Path
}

# --- Load minimal functions needed for queue tests ---
. (Join-Path $repoRoot 'Source\Private\Use-FileMutex.ps1')
. (Join-Path $repoRoot 'Source\Public\Queue\Add-QueueItemFolder.ps1')
. (Join-Path $repoRoot 'Source\Public\Queue\Get-NextQueueItemFolder.ps1')
. (Join-Path $repoRoot 'Source\Public\Queue\Get-QueueItemsFolder.ps1')

# Isolate queue folder for this test run
$testQueue = Join-Path $env:TEMP ("DoclingQueue_Test_" + [guid]::NewGuid().ToString('N'))
$env:TEMP_BACKUP_FOR_TEST = $env:TEMP

# Monkey-patch by temporarily pointing TEMP so queue functions use our folder.
# The functions hardcode "$env:TEMP\DoclingQueue" - use a unique subfolder by
# setting TEMP to our test parent so DoclingQueue lands under it.
$isolatedTemp = Join-Path $env:TEMP ("PSDoclingAudit_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $isolatedTemp -Force | Out-Null
$originalTemp = $env:TEMP
$env:TEMP = $isolatedTemp
$queueDir = Join-Path $env:TEMP 'DoclingQueue'

try {
    # Test: add then get returns same id (FIFO)
    $id1 = [guid]::NewGuid().ToString()
    $id2 = [guid]::NewGuid().ToString()
    Add-QueueItemFolder $id1 | Out-Null
    Add-QueueItemFolder $id2 | Out-Null
    $listed = @(Get-QueueItemsFolder)
    Assert-True ($listed.Count -eq 2) "Get-QueueItemsFolder lists 2 after two adds (got $($listed.Count))"

    $next1 = Get-NextQueueItemFolder
    Assert-True ($next1 -eq $id1) "Get-NextQueueItemFolder returns oldest first (expected $id1 got $next1)"

    $next2 = Get-NextQueueItemFolder
    Assert-True ($next2 -eq $id2) "Get-NextQueueItemFolder returns second item (expected $id2 got $next2)"

    $next3 = Get-NextQueueItemFolder
    Assert-True ($null -eq $next3) "Get-NextQueueItemFolder returns null when empty"

    $remaining = @(Get-QueueItemsFolder)
    Assert-True ($remaining.Count -eq 0) "Queue folder empty after claims"

    # Test: claim removes the file (no double-claim of same file)
    $id3 = [guid]::NewGuid().ToString()
    Add-QueueItemFolder $id3 | Out-Null
    $claimed = Get-NextQueueItemFolder
    $filesLeft = @(Get-ChildItem -Path $queueDir -Filter '*.queue' -ErrorAction SilentlyContinue)
    Assert-True ($claimed -eq $id3) "Single-item claim returns correct id"
    Assert-True ($filesLeft.Count -eq 0) "Claim deletes queue file (no leftover .queue)"
}
finally {
    $env:TEMP = $originalTemp
    if (Test-Path $isolatedTemp) {
        Remove-Item $isolatedTemp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# --- Static wiring checks (source of truth) ---
$apiServer = Get-Content (Join-Path $repoRoot 'Source\Public\Services\Start-APIServer.ps1') -Raw
Assert-True ($apiServer -match 'Add-QueueItemFolder\s+\$documentId') "Reprocess path enqueues via Add-QueueItemFolder"
Assert-True ($apiServer -notmatch 'Add-QueueItem\s+\$reprocessItem') "Reprocess path no longer uses Add-QueueItem `$reprocessItem"

$statusFn = Get-Content (Join-Path $repoRoot 'Source\Public\System\Get-DoclingSystemStatus.ps1') -Raw
Assert-True ($statusFn -match 'Get-QueueItemsFolder') "Get-DoclingSystemStatus uses folder queue"
Assert-True ($statusFn -notmatch 'Get-QueueItems(?!Folder)') "Get-DoclingSystemStatus does not use JSON Get-QueueItems"

$frontend = Get-Content (Join-Path $repoRoot 'Source\Public\Services\New-FrontendFiles.ps1') -Raw
# Fixed pattern uses JS concat: downloadDocument(\'' + id + '\')
Assert-True ($frontend -match "downloadDocument\(\\'' \+ id \+ '\\'\)") "Frontend downloadDocument onclick uses proper JS concatenation"
Assert-True ($frontend -notmatch 'downloadDocument\(\\''" \+ id \+ "\\''\)') "Frontend downloadDocument onclick no longer embeds literal + id +"

Write-Host ""
Write-Host "Results: $passed passed, $failed failed"
if ($failed -gt 0) { exit 1 } else { exit 0 }
