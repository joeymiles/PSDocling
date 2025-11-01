param(
  [switch]$GenerateFrontend,
  [switch]$OpenBrowser,
  [switch]$UseWebView,
  [switch]$SkipPythonCheck,
  [switch]$EnsureUrlAcl,
  [switch]$ClearHistory,
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
  # Try to use built module first, fall back to source
  $buildModulePath = Join-Path $PSScriptRoot 'Build\PSDocling.psm1'
  $sourceModulePath = Join-Path $PSScriptRoot 'PSDocling.psm1'

  if (Test-Path $buildModulePath) {
    $modulePath = $buildModulePath
    Write-Info "Using built module from Build folder"
  } elseif (Test-Path $sourceModulePath) {
    $modulePath = $sourceModulePath
    Write-Warn "Build folder not found, using source module"
  } else {
    Write-Err "PSDocling.psm1 not found in Build or root folder"
    exit 1
  }

  Write-Info "Importing module..."
  Import-Module $modulePath -Force

  if ($EnsureUrlAcl) {
    Ensure-UrlAcl -port $ApiPort
  }

  Write-Info "Initializing system..."
  $initParams = @{}
  if ($GenerateFrontend) { $initParams['GenerateFrontend'] = $true }
  if ($SkipPythonCheck) { $initParams['SkipPythonCheck'] = $true }
  if ($ClearHistory) { $initParams['ClearHistory'] = $true }

  Initialize-DoclingSystem @initParams | Out-Null

  # Try to apply port overrides to the exported configuration hashtable
  try {
    if ($script:DoclingSystem) { } # no-op; keeps analyzer happy
  } catch { }

  if (Get-Variable -Name DoclingSystem -Scope Global -ErrorAction SilentlyContinue) {
    if ($null -ne $DoclingSystem.Backend) { $DoclingSystem.Backend.APIPort = $ApiPort }
    if ($null -ne $DoclingSystem.Frontend) { $DoclingSystem.Frontend.WebServerPort = $WebPort }
  }

  Write-Info "Starting services..."
  $startParams = @{}
  if ($OpenBrowser) { $startParams['OpenBrowser'] = $true }
  if ($UseWebView) { $startParams['UseWebView'] = $true }

  Start-DoclingSystem @startParams | Out-Null

  Write-Ok "Backend API running on http://localhost:$ApiPort"

  if ($UseWebView) {
    Write-Info "Native window launched with PyWebView"
  } elseif ($OpenBrowser) {
    Write-Info "Browser opened at http://localhost:$WebPort"
  } else {
    Write-Info "Tip: Use -OpenBrowser to open in browser, or -UseWebView for native window"
    Write-Info "     Frontend available at: http://localhost:$WebPort"
  }

} finally {
  Pop-Location
}

