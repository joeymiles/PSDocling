@{
    # Module manifest for PSDocling
    RootModule = 'PSDocling.psm1'
    ModuleVersion = '2.1.1'

    # Unique identifier for this module
    GUID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'

    # Author and company information
    Author = 'Joey A Miles'
    CompanyName = 'Just A Guy Doing Cool Stuff'
    Copyright = '(c) 2025. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'PowerShell-based document processing system that converts various document formats (PDF, DOCX, PPTX, XLSX, HTML, MD) to Markdown, JSON, HTML, Doctags using the Python Docling library. Includes REST API server, document processor, and web frontend.'

    # Minimum version of PowerShell required
    PowerShellVersion = '5.1'

    # Required .NET Framework version
    DotNetFrameworkVersion = '4.7.2'

    # Compatible PowerShell editions
    CompatiblePSEditions = @('Desktop', 'Core')

    # Functions to export from this module
    FunctionsToExport = @(
        'Initialize-DoclingSystem',
        'Start-DoclingSystem',
        'Add-DocumentToQueue',
        'Start-DocumentProcessor',
        'Start-APIServer',
        'New-FrontendFiles',
        'Get-DoclingSystemStatus',
        'Get-PythonStatus',
        'Get-QueueItems',
        'Set-QueueItems',
        'Add-QueueItem',
        'Get-NextQueueItem',
        'Get-ProcessingStatus',
        'Set-ProcessingStatus',
        'Update-ItemStatus'
    )

    # Cmdlets to export from this module
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module
    AliasesToExport = @()

    # Private data to pass to the module
    PrivateData = @{
        PSData = @{
            # Tags applied to this module
            Tags = @('Document', 'Processing', 'Conversion', 'PDF', 'Markdown', 'Docling', 'REST', 'API', 'Web')

            # A URL to the license for this module
            LicenseUri = ''

            # A URL to the main website for this project
            ProjectUri = ''

            # A URL to an icon representing this module
            IconUri = ''

            # Release notes for this version
            ReleaseNotes = @"
Version 2.1.1:
- Queue-based document processing architecture
- Multi-process system (API server, processor, web frontend)
- Support for PDF, DOCX, PPTX, XLSX, HTML, MD, and image formats
- REST API with comprehensive error handling
- Web frontend with drag-drop file upload
- Python Docling library integration
- Cross-platform PowerShell Core support
"@

            # External modules that this module depends on
            ExternalModuleDependencies = @()
        }
    }

    # Help Info URI for this module
    HelpInfoURI = ''

    # Default prefix for commands exported from this module
    DefaultCommandPrefix = ''

    # File list for this module
    FileList = @(
        'PSDocling.psm1',
        'PSDocling.psd1',
        'Start-All.ps1',
        'Stop-All.ps1',
        'HowTo.ps1'
    )

    # Required modules that must be imported
    RequiredModules = @()

    # Assemblies that must be loaded
    RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment
    ScriptsToProcess = @()

    # Type files (.ps1xml) to be loaded
    TypesToProcess = @()

    # Format files (.ps1xml) to be loaded
    FormatsToProcess = @()

    # Modules to import as nested modules
    NestedModules = @()
}