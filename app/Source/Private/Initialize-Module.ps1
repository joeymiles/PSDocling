<#
.SYNOPSIS
    Module initialization code for PSDocling
.DESCRIPTION
    Contains module-scoped variables and initialization logic
#>

# Docling Document Processing System
# Version: 3.3.2

# Runs before private functions are defined in the built module, so the
# home is resolved inline here. Keep in sync with Get-DoclingPath.
$doclingHome = if ($env:PSDOCLING_HOME) { $env:PSDOCLING_HOME } else { Join-Path $env:LOCALAPPDATA 'PSDocling' }
$doclingData = Join-Path $doclingHome 'data'

$script:DoclingSystem = @{
    Version          = "3.3.2"
    ModulePath       = $PSCommandPath
    HomeDirectory    = $doclingHome
    TempDirectory    = Join-Path $doclingData 'uploads'
    OutputDirectory  = Join-Path $doclingData 'output'
    APIPort          = 8080
    WebPort          = 8081
    QueueFile        = Join-Path $doclingData 'queue.json'
    StatusFile       = Join-Path $doclingData 'status.json'
    PythonAvailable  = $false
    ProcessingStatus = @{}
}

# Function to check and install required Python packages

