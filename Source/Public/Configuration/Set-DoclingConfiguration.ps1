function Set-DoclingConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Settings,

        [string]$Path = "$PSScriptRoot\..\..\..\PSDocling.config.psd1",

        [switch]$Merge
    )

    try {
        if ($Merge -and (Test-Path $Path)) {
            $existing = Import-PowerShellDataFile -Path $Path

            # Deep merge the settings
            foreach ($key in $Settings.Keys) {
                if ($existing.ContainsKey($key) -and $existing[$key] -is [hashtable] -and $Settings[$key] -is [hashtable]) {
                    foreach ($subkey in $Settings[$key].Keys) {
                        $existing[$key][$subkey] = $Settings[$key][$subkey]
                    }
                } else {
                    $existing[$key] = $Settings[$key]
                }
            }
            $Settings = $existing
        }

        # Export to file
        $content = ConvertTo-PSD -InputObject $Settings
        Set-Content -Path $Path -Value $content -Encoding UTF8

        Write-Verbose "Configuration saved to $Path"
        return $true
    }
    catch {
        Write-Error "Failed to save configuration: $_"
        return $false
    }
}

function ConvertTo-PSD {
    param($InputObject, $Depth = 0)

    $indent = "    " * $Depth

    if ($InputObject -is [hashtable]) {
        $lines = @("@{")
        foreach ($key in $InputObject.Keys) {
            $value = ConvertTo-PSD -InputObject $InputObject[$key] -Depth ($Depth + 1)
            $lines += "$indent    $key = $value"
        }
        $lines += "$indent}"
        return ($lines -join "`n")
    }
    elseif ($InputObject -is [array]) {
        $items = $InputObject | ForEach-Object {
            ConvertTo-PSD -InputObject $_ -Depth ($Depth + 1)
        }
        return "@(" + ($items -join ", ") + ")"
    }
    elseif ($InputObject -is [string]) {
        return "'$($InputObject -replace "'", "''")'"
    }
    elseif ($InputObject -is [bool]) {
        return if ($InputObject) { '$true' } else { '$false' }
    }
    elseif ($null -eq $InputObject) {
        return '$null'
    }
    else {
        return $InputObject.ToString()
    }
}