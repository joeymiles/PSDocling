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

$RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $RepoRoot 'Source')) -and -not (Test-Path (Join-Path $RepoRoot 'Build'))) {
  $RepoRoot = $PSScriptRoot
}

Push-Location $RepoRoot
try {
  # Prefer installed module, then Build/, then fail with install hint
  $buildModulePath = Join-Path $RepoRoot 'Build\PSDocling.psm1'
  $installed = Get-Module -ListAvailable PSDocling -ErrorAction SilentlyContinue | Select-Object -First 1

  if ($installed) {
    Write-Info "Using installed module: $($installed.Path)"
    Import-Module PSDocling -Force
  } elseif (Test-Path $buildModulePath) {
    Write-Info "Using built module from Build/"
    Import-Module $buildModulePath -Force
  } else {
    Write-Err "PSDocling not found. Install first: .\scripts\Install-DoclingModule.ps1"
    exit 1
  }

  if ($EnsureUrlAcl) {
    Ensure-UrlAcl -port $ApiPort
  }

  Write-Info "Initializing system..."
  $initParams = @{}
  if ($GenerateFrontend) { $initParams['GenerateFrontend'] = $true }
  if ($SkipPythonCheck) { $initParams['SkipPythonCheck'] = $true }
  if ($ClearHistory) { $initParams['ClearHistory'] = $true }

  Initialize-DoclingSystem @initParams | Out-Null

  if (Get-Variable -Name DoclingSystem -Scope Global -ErrorAction SilentlyContinue) {
    if ($null -ne $DoclingSystem.Backend) { $DoclingSystem.Backend.APIPort = $ApiPort }
    if ($null -ne $DoclingSystem.Frontend) { $DoclingSystem.Frontend.WebServerPort = $WebPort }
  }

  # Also try script-scoped ports used by current module
  try {
    $mod = Get-Module PSDocling
    if ($mod) {
      # Module uses $script:DoclingSystem.APIPort / WebPort — set via Initialize defaults;
      # Start-All historically overrode global. Re-init ports by direct call if helpers exist.
    }
  } catch { }

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
