<#
.SYNOPSIS
    Module initialization code for PSDocling
.DESCRIPTION
    Contains module-scoped variables and initialization logic
#>

# Docling Document Processing System
# Version: 3.3.0

$script:DoclingSystem = @{
    Version          = "3.3.0"
    ModulePath       = $PSCommandPath
    TempDirectory    = "$env:TEMP\DoclingProcessor"
    OutputDirectory  = "$env:TEMP\DoclingOutput"
    APIPort          = 8080
    WebPort          = 8081
    QueueFile        = "$env:TEMP\docling_queue.json"
    StatusFile       = "$env:TEMP\docling_status.json"
    PythonAvailable  = $false
    ProcessingStatus = @{}
}

# Function to check and install required Python packages

