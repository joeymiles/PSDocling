function Get-DoclingConfiguration {
    [CmdletBinding()]
    param(
        [string]$Path,
        [string]$Section
    )

    try {
        # Build search paths based on whether Path was provided
        if ($Path) {
            # If explicit path provided, use it first
            $searchPaths = @($Path)
        } else {
            # Default search locations
            $searchPaths = @(
                (Join-Path (Get-Location) "PSDocling.config.psd1")
            )

            # Add user-specific module paths
            $userModulePath = Join-Path $env:USERPROFILE "Documents\Scripts\GitHub\Docling4\PSDocling.config.psd1"
            if (Test-Path $userModulePath) {
                $searchPaths += $userModulePath
            }

            # Add script-relative paths if running from module
            if ($PSScriptRoot) {
                $searchPaths += @(
                    (Join-Path $PSScriptRoot "PSDocling.config.psd1"),
                    (Join-Path (Split-Path $PSScriptRoot -Parent) "PSDocling.config.psd1"),
                    (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "PSDocling.config.psd1"),
                    (Join-Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) "PSDocling.config.psd1")
                )
            }
        }

        $configFile = $null
        foreach ($searchPath in $searchPaths) {
            if ($searchPath -and (Test-Path $searchPath)) {
                $configFile = $searchPath
                Write-Verbose "Found configuration at: $configFile"
                break
            }
        }

        if ($configFile) {
            # Use Import-PowerShellDataFile if available (PS 5.0+), otherwise use Invoke-Expression
            if (Get-Command Import-PowerShellDataFile -ErrorAction SilentlyContinue) {
                $config = Import-PowerShellDataFile -Path $configFile
            } else {
                # Fallback for older PowerShell versions
                $content = Get-Content -Path $configFile -Raw
                $config = Invoke-Expression $content
            }
        } else {
            Write-Verbose "Configuration file not found, using defaults"
            $config = @{
                Server = @{
                    APIHost = "localhost"
                    APIPort = 8080
                    WebHost = "localhost"
                    WebPort = 8081
                    Protocol = "http"
                }
                Endpoints = @{
                    APIBaseURL = $null
                }
            }
        }

        # Post-process the configuration to expand paths and convert types
        if ($config) {
            # Expand TEMP paths in Processing section
            if ($config.Processing) {
                foreach ($key in @('TempDirectory', 'QueueFile', 'StatusFile')) {
                    if ($config.Processing.ContainsKey($key) -and $config.Processing[$key] -like 'TEMP\*') {
                        $config.Processing[$key] = $config.Processing[$key] -replace '^TEMP\\', "$env:TEMP\"
                    }
                }

                # Convert string booleans to actual booleans in Processing
                foreach ($key in @('EnableCodeEnrichment', 'EnableFormulaEnrichment', 'EnablePictureClassification')) {
                    if ($config.Processing.ContainsKey($key)) {
                        if ($config.Processing[$key] -eq 'true') { $config.Processing[$key] = $true }
                        elseif ($config.Processing[$key] -eq 'false') { $config.Processing[$key] = $false }
                    }
                }
            }

            # Convert string booleans in other sections
            if ($config.Server -and $config.Server.ContainsKey('EnableCORS')) {
                if ($config.Server.EnableCORS -eq 'true') { $config.Server.EnableCORS = $true }
                elseif ($config.Server.EnableCORS -eq 'false') { $config.Server.EnableCORS = $false }
            }

            if ($config.Frontend -and $config.Frontend.ContainsKey('EnableFilePreview')) {
                if ($config.Frontend.EnableFilePreview -eq 'true') { $config.Frontend.EnableFilePreview = $true }
                elseif ($config.Frontend.EnableFilePreview -eq 'false') { $config.Frontend.EnableFilePreview = $false }
            }

            if ($config.Deployment) {
                foreach ($key in @('AllowRemoteConnections', 'RequireAuthentication')) {
                    if ($config.Deployment.ContainsKey($key)) {
                        if ($config.Deployment[$key] -eq 'true') { $config.Deployment[$key] = $true }
                        elseif ($config.Deployment[$key] -eq 'false') { $config.Deployment[$key] = $false }
                    }
                }
            }
        }

        if ($Section) {
            return $config.$Section
        }

        # Ensure Endpoints section exists
        if (-not $config.ContainsKey('Endpoints')) {
            $config.Endpoints = @{}
        }

        # Build computed properties
        if (-not $config.Endpoints.APIBaseURL) {
            $server = $config.Server
            $config.Endpoints.APIBaseURL = "$($server.Protocol)://$($server.APIHost):$($server.APIPort)"
        }

        return $config
    }
    catch {
        Write-Error "Failed to load configuration: $_"
        return $null
    }
}