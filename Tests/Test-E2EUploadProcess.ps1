#Requires -Version 5.1
<#
.SYNOPSIS
    Minimal E2E smoke: upload -> start-conversion -> Completed -> download/result.
.DESCRIPTION
    Requires a running PSDocling stack (API :8080). Creates Tests/fixtures/tiny.pdf if missing.
    Exit 0 on success, 1 on failure.
#>
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$fixtureDir = Join-Path $PSScriptRoot 'fixtures'
$pdfPath = Join-Path $fixtureDir 'tiny.pdf'
$api = 'http://localhost:8080'

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "FAIL: $Message" }
    Write-Host "PASS: $Message" -ForegroundColor Green
}

# Health
try {
    $health = Invoke-RestMethod "$api/api/health" -TimeoutSec 5
    Assert-True ($health.status -eq 'healthy') "API healthy"
} catch {
    throw "API not reachable at $api. Start with: .\scripts\Start-All.ps1 -OpenBrowser"
}

# Tiny PDF fixture
if (-not (Test-Path $pdfPath)) {
    New-Item -ItemType Directory -Path $fixtureDir -Force | Out-Null
    $pdf = @"
%PDF-1.1
1 0 obj<< /Type /Catalog /Pages 2 0 R >>endobj
2 0 obj<< /Type /Pages /Kids [3 0 R] /Count 1 >>endobj
3 0 obj<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 144] /Contents 4 0 R /Resources<< /Font<< /F1 5 0 R >> >> >>endobj
4 0 obj<< /Length 44 >>stream
BT /F1 24 Tf 50 80 Td (Hello PSDocling) Tj ET
endstream endobj
5 0 obj<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000266 00000 n 
0000000360 00000 n 
trailer<< /Size 6 /Root 1 0 R >>
startxref
441
%%EOF
"@
    [IO.File]::WriteAllBytes($pdfPath, [Text.Encoding]::ASCII.GetBytes($pdf))
}

$bytes = [IO.File]::ReadAllBytes($pdfPath)
$b64 = [Convert]::ToBase64String($bytes)
$upload = Invoke-RestMethod -Uri "$api/api/upload" -Method POST -ContentType 'application/json' -Body (@{
    fileName = 'tiny.pdf'
    dataBase64 = $b64
    exportFormat = 'markdown'
} | ConvertTo-Json)
Assert-True ($upload.success -eq $true -and $upload.documentId) "Upload returns documentId"
$docId = $upload.documentId

$docs = Invoke-RestMethod "$api/api/documents" -TimeoutSec 10
Assert-True (@($docs).Count -ge 1) "Documents API lists uploaded item"

$start = Invoke-RestMethod -Uri "$api/api/start-conversion" -Method POST -ContentType 'application/json' -Body (@{
    documentId = $docId
    exportFormat = 'markdown'
} | ConvertTo-Json)
Assert-True ($start.success -eq $true) "Start-conversion accepted"

$deadline = (Get-Date).AddMinutes(6)
$status = $null
do {
    Start-Sleep -Seconds 2
    $all = Get-Content "$env:TEMP\docling_status.json" -Raw | ConvertFrom-Json
    $status = $all.$docId
    if ($null -eq $status) { continue }
    Write-Host "  status=$($status.Status) progress=$($status.Progress)"
} while ($status.Status -notin @('Completed','Failed','Error','Cancelled') -and (Get-Date) -lt $deadline)

Assert-True ($status.Status -eq 'Completed') "Document reached Completed (got $($status.Status))"
Assert-True ($status.OutputFile -and (Test-Path $status.OutputFile)) "Output file exists: $($status.OutputFile)"

$result = Invoke-WebRequest -Uri "$api/api/result/$docId" -UseBasicParsing -TimeoutSec 15
Assert-True ($result.StatusCode -eq 200 -and $result.RawContentLength -gt 0) "Result download HTTP 200 with body"

$dl = Invoke-WebRequest -Uri "$api/api/download/$docId" -UseBasicParsing -TimeoutSec 15
Assert-True ($dl.StatusCode -eq 200 -and $dl.RawContentLength -gt 0) "Download endpoint HTTP 200 with body"

Write-Host "`nE2E smoke passed for $docId" -ForegroundColor Green

