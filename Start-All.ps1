param(
  [switch]$GenerateFrontend,
  [switch]$OpenBrowser,
  [switch]$SkipPythonCheck,
  [switch]$EnsureUrlAcl,
  [int]$ApiPort = 8080,
  [int]$WebPort = 8081
)

function Write-Info($msg)  { Write-Host $msg -ForegroundColor Cyan }
function Write-Ok($msg)    { Write-Host $msg -ForegroundColor Green }
function Write-Warn($msg)  { Write-Host $msg -ForegroundColor Yellow }
function Write-Err($msg)   { Write-Host $msg -ForegroundColor Red }

function Test-IsAdmin {
  $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-UrlAcl($port) {
  if (-not (Test-IsAdmin)) {
    Write-Warn "Skipping URL ACL for port $port (admin required). Run PowerShell as Administrator to add it."
    return
  }
  $urls = @("http://localhost:$port/", "http://127.0.0.1:$port/")
  foreach ($u in $urls) {
    try {
      Write-Info "Adding URL ACL: $u"
      & netsh http add urlacl url=$u user=Everyone | Out-Null
    } catch {
      Write-Warning "URL ACL add failed or already exists for: $u $($_.Exception.Message)"
    }
  }
}

Push-Location $PSScriptRoot
try {
  $modulePath = Join-Path $PSScriptRoot 'PSDocling.psm1'
  if (-not (Test-Path $modulePath)) {
    Write-Err "PSDocling.psm1 not found at $modulePath"
    exit 1
  }

  Write-Info "Importing module..."
  Import-Module $modulePath -Force

  if ($EnsureUrlAcl) {
    Ensure-UrlAcl -port $ApiPort
  }

  Write-Info "Initializing system..."
  if ($GenerateFrontend -and $SkipPythonCheck) {
    Initialize-DoclingSystem -GenerateFrontend -SkipPythonCheck | Out-Null
  } elseif ($GenerateFrontend) {
    Initialize-DoclingSystem -GenerateFrontend | Out-Null
  } elseif ($SkipPythonCheck) {
    Initialize-DoclingSystem -SkipPythonCheck | Out-Null
  } else {
    Initialize-DoclingSystem | Out-Null
  }

  # Try to apply port overrides to the exported configuration hashtable
  try {
    if ($script:DoclingSystem) { } # no-op; keeps analyzer happy
  } catch { }

  if (Get-Variable -Name DoclingSystem -Scope Global -ErrorAction SilentlyContinue) {
    if ($null -ne $DoclingSystem.Backend) { $DoclingSystem.Backend.APIPort = $ApiPort }
    if ($null -ne $DoclingSystem.Frontend) { $DoclingSystem.Frontend.WebServerPort = $WebPort }
  }

  Write-Info "Starting services..."
  if ($OpenBrowser) {
    Start-DoclingSystem -OpenBrowser | Out-Null
  } else {
    Start-DoclingSystem | Out-Null
  }

  Write-Ok "Backend API:   http://localhost:$ApiPort"
  Write-Ok "Frontend UI:   http://localhost:$WebPort"
  Write-Info "Tip: Use -OpenBrowser to open the UI automatically."

} finally {
  Pop-Location
}

