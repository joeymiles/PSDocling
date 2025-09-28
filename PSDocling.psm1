# Docling Document Processing System
# Version: 2.1.2

$script:DoclingSystem = @{
    Version          = "2.1.2"
    TempDirectory    = "$env:TEMP\DoclingProcessor"
    OutputDirectory  = ".\ProcessedDocuments"
    APIPort          = 8080
    WebPort          = 8081
    QueueFile        = "$env:TEMP\docling_queue.json"
    StatusFile       = "$env:TEMP\docling_status.json"
    PythonAvailable  = $false
    ProcessingStatus = @{}
}

# Auto-detect Python on module load
try {
    $pythonCmd = Get-Command python -ErrorAction Stop
    if ($pythonCmd) {
        $script:DoclingSystem.PythonAvailable = $true
        Write-Host "Python auto-detected" -ForegroundColor Green
    }
}
catch {
    Write-Host "Python not found during auto-detection" -ForegroundColor Yellow
}

# File-based queue management
function Get-QueueItems {
    if (Test-Path $script:DoclingSystem.QueueFile) {
        try {
            $content = Get-Content $script:DoclingSystem.QueueFile -Raw
            if ($content.Trim() -eq "[]") {
                return @()
            }
            # Force array conversion in PowerShell 5.1
            $items = @($content | ConvertFrom-Json)
            return $items
        }
        catch {
            return @()
        }
    }
    return @()
}

function Set-QueueItems {
    param([array]$Items = @())
    # Ensure we always store as a JSON array, even for single items
    if ($Items.Count -eq 0) {
        "[]" | Set-Content $script:DoclingSystem.QueueFile -Encoding UTF8
    }
    elseif ($Items.Count -eq 1) {
        "[" + ($Items[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $script:DoclingSystem.QueueFile -Encoding UTF8
    }
    else {
        $Items | ConvertTo-Json -Depth 10 | Set-Content $script:DoclingSystem.QueueFile -Encoding UTF8
    }
}

function Add-QueueItem {
    param($Item)
    $queue = Get-QueueItems
    # Force array addition
    $newQueue = @($queue) + @($Item)
    Set-QueueItems $newQueue
}

function Get-NextQueueItem {
    $queue = Get-QueueItems
    if ($queue.Count -gt 0) {
        $item = $queue[0]
        $remaining = if ($queue.Count -gt 1) { $queue[1..($queue.Count - 1)] } else { @() }
        Set-QueueItems $remaining
        return $item
    }
    return $null
}

# File-based status management for cross-process sharing
function Get-ProcessingStatus {
    if (Test-Path $script:DoclingSystem.StatusFile) {
        try {
            $content = Get-Content $script:DoclingSystem.StatusFile -Raw
            $jsonObj = $content | ConvertFrom-Json

            # Convert PSCustomObject to hashtable manually
            $hashtable = @{}
            $jsonObj.PSObject.Properties | ForEach-Object {
                $hashtable[$_.Name] = $_.Value
            }
            return $hashtable
        }
        catch {
            return @{}
        }
    }
    return @{}
}

function Set-ProcessingStatus {
    param([hashtable]$Status)
    $Status | ConvertTo-Json -Depth 10 | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8
}

function Update-ItemStatus {
    param($Id, $Updates)
    $status = Get-ProcessingStatus

    # Convert existing item to hashtable if it's a PSObject
    if ($status[$Id]) {
        if ($status[$Id] -is [PSCustomObject]) {
            $itemHash = @{}
            $status[$Id].PSObject.Properties | ForEach-Object {
                $itemHash[$_.Name] = $_.Value
            }
            $status[$Id] = $itemHash
        }
    }
    else {
        $status[$Id] = @{}
    }

    # Apply updates
    foreach ($key in $Updates.Keys) {
        $status[$Id][$key] = $Updates[$key]
    }

    Set-ProcessingStatus $status
    # Also update local cache
    $script:DoclingSystem.ProcessingStatus[$Id] = $status[$Id]
}

function Test-SecureFileName {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,
        [string[]]$AllowedExtensions = @('.pdf', '.docx', '.xlsx', '.pptx', '.md', '.html', '.xhtml', '.csv', '.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.webp', '.txt')
    )

    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return $false
    }

    # Extract just the base filename to prevent directory traversal
    $safeFileName = [System.IO.Path]::GetFileName($FileName)

    # Additional security checks
    if ($safeFileName -ne $FileName) {
        # Original contained path separators - potential traversal attempt
        return $false
    }

    # Check for invalid characters (beyond what GetFileName handles)
    if ($safeFileName -match '[<>:"|?*]') {
        return $false
    }

    # Check length limits (NTFS limit is 255, but we'll be more conservative)
    if ($safeFileName.Length -gt 200) {
        return $false
    }

    # Check extension if provided
    if ($AllowedExtensions.Count -gt 0) {
        $extension = [System.IO.Path]::GetExtension($safeFileName).ToLower()
        if ($extension -notin $AllowedExtensions) {
            return $false
        }
    }

    # Check for reserved names (Windows)
    $reservedNames = @('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($safeFileName).ToUpper()
    if ($nameWithoutExt -in $reservedNames) {
        return $false
    }

    return $true
}

function Get-SecureFileName {
    param(
        [Parameter(Mandatory)]
        [string]$FileName,
        [string[]]$AllowedExtensions = @('.pdf', '.docx', '.xlsx', '.pptx', '.md', '.html', '.xhtml', '.csv', '.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.webp', '.txt')
    )

    if (-not (Test-SecureFileName -FileName $FileName -AllowedExtensions $AllowedExtensions)) {
        throw "Invalid or potentially dangerous filename: $FileName"
    }

    return [System.IO.Path]::GetFileName($FileName)
}

<#
.SYNOPSIS
    Initializes the Docling document processing system.

.DESCRIPTION
    Sets up the PSDocling system by creating necessary directories, initializing the queue system, checking Python and Docling dependencies, and optionally generating the web frontend files. This function must be run before using other Docling functions.

.PARAMETER SkipPythonCheck
    Skips the Python and Docling dependency verification. Useful for development or when running in simulation mode.

.PARAMETER GenerateFrontend
    Generates the web frontend files after initialization. The frontend provides a web interface for document upload and processing.

.EXAMPLE
    Initialize-DoclingSystem

    Basic initialization with Python dependency checking.

.EXAMPLE
    Initialize-DoclingSystem -GenerateFrontend

    Initialize the system and generate the web frontend files.

.EXAMPLE
    Initialize-DoclingSystem -SkipPythonCheck -GenerateFrontend

    Initialize in development mode without Python checking and generate frontend files.

.NOTES
    Creates the following directories:
    - $env:TEMP\DoclingProcessor (temporary processing directory)
    - .\ProcessedDocuments (output directory for converted documents)

    Initializes the following files:
    - $env:TEMP\docling_queue.json (processing queue)
    - $env:TEMP\docling_status.json (status tracking)

.LINK
    https://github.com/DS4SD/docling
#>
<#
.SYNOPSIS
    Initializes the PSDocling document processing system.

.DESCRIPTION
    Sets up the required directories, initializes the queue and status files, checks for Python and Docling library availability, and optionally generates the web frontend.

.PARAMETER SkipPythonCheck
    Skip checking for Python and Docling library availability during initialization.

.PARAMETER GenerateFrontend
    Generate the web frontend files after initialization.

.EXAMPLE
    Initialize-DoclingSystem
    Initializes the system with Python checks.

.EXAMPLE
    Initialize-DoclingSystem -SkipPythonCheck -GenerateFrontend
    Initializes the system without Python checks and generates the web frontend.

.NOTES
    This function must be called before using other PSDocling functions. It creates the necessary temporary directories and queue files for the document processing system.
#>
function Initialize-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$SkipPythonCheck,
        [switch]$GenerateFrontend
    )

    Write-Host "Initializing PS Docling System v$($script:DoclingSystem.Version)" -ForegroundColor Cyan

    # Create directories
    @($script:DoclingSystem.TempDirectory, $script:DoclingSystem.OutputDirectory) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            Write-Host "Created directory: $_" -ForegroundColor Green
        }
    }

    # Initialize queue
    Set-QueueItems @()
    @{} | ConvertTo-Json | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8

    # Check Python
    if (-not $SkipPythonCheck) {
        try {
            $version = & python --version 2>&1
            if ($version -match "Python") {
                Write-Host "Python found: $version" -ForegroundColor Green
                $script:DoclingSystem.PythonAvailable = $true

                # Check Docling
                $pipOutput = & python -m pip show docling 2>&1
                if (-not ($pipOutput -match "Name: docling")) {
                    Write-Host "Installing Docling..." -ForegroundColor Yellow
                    & python -m pip install docling --quiet
                }
                Write-Host "Docling ready" -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Python not found - using simulation mode"
        }
    }

    if ($GenerateFrontend) {
        New-FrontendFiles
    }

    Write-Host "System initialized" -ForegroundColor Green
}

<#
.SYNOPSIS
    Adds documents to the Docling processing queue.

.DESCRIPTION
    Queues one or more documents for processing with the Docling system. Validates file formats and creates processing items with unique IDs. Supported formats include PDF, DOCX, XLSX, PPTX, MD, HTML, CSV, and various image formats.

.PARAMETER Path
    One or more file paths to add to the processing queue. Accepts pipeline input.

.PARAMETER ExportFormat
    The output format for document conversion. Valid values: 'markdown', 'json', 'html', 'doctags'. Default is 'markdown'.

.PARAMETER EmbedImages
    When specified, embeds images directly in the output format (where supported).

.EXAMPLE
    Add-DocumentToQueue -Path "C:\Documents\report.pdf"

    Adds a PDF document to the processing queue with default markdown output.

.EXAMPLE
    Add-DocumentToQueue -Path "C:\Documents\report.pdf" -ExportFormat "json" -EmbedImages

    Adds a PDF document to the queue for JSON conversion with embedded images.

.EXAMPLE
    Get-ChildItem "C:\Documents\*.pdf" | Add-DocumentToQueue -ExportFormat "html"

    Adds all PDF files from a directory to the processing queue for HTML conversion.

.OUTPUTS
    String. Returns the unique document ID for each queued document.

.NOTES
    Supported file formats:
    - Documents: PDF, DOCX, XLSX, PPTX, MD, HTML, XHTML, CSV
    - Images: PNG, JPG, JPEG, TIFF, TIF, BMP, WEBP

    The function validates file existence and format before queuing.
    Each document receives a unique GUID identifier for tracking.
#>
<#
.SYNOPSIS
    Adds one or more documents to the processing queue.

.DESCRIPTION
    Queues documents for processing by the PSDocling system. Supports various document formats including PDF, DOCX, XLSX, PPTX, MD, HTML, CSV, and image formats.

.PARAMETER Path
    The path(s) to the document(s) to be processed. Accepts pipeline input and multiple values.

.PARAMETER ExportFormat
    The output format for the processed document. Default is 'markdown'. Supported formats: markdown, json, html, doctags.

.PARAMETER EmbedImages
    When specified, embeds images directly in the output format (where supported).

.PARAMETER EnrichCode
    When specified, enables code understanding enrichment to parse code blocks and identify programming languages.

.PARAMETER EnrichFormula
    When specified, enables formula understanding enrichment to extract LaTeX representations from mathematical formulas.

.PARAMETER EnrichPictureClasses
    When specified, enables picture classification enrichment to identify chart types, diagrams, logos, and signatures.

.PARAMETER EnrichPictureDescription
    When specified, enables picture description enrichment using Granite Vision model to generate descriptive text for images.

.EXAMPLE
    Add-DocumentToQueue -Path "C:\Documents\report.pdf"
    Adds a PDF document to the processing queue with default markdown output.

.EXAMPLE
    Add-DocumentToQueue -Path "C:\Documents\report.pdf" -ExportFormat "json" -EmbedImages
    Adds a PDF document to the queue for JSON output with embedded images.

.EXAMPLE
    Add-DocumentToQueue -Path "C:\Documents\technical.pdf" -EnrichCode -EnrichFormula
    Adds a PDF with code and formula enrichment enabled.

.EXAMPLE
    Add-DocumentToQueue -Path "C:\Documents\research.pdf" -EnrichPictureDescription -EnrichPictureClasses
    Adds a PDF with picture analysis and description using Granite Vision model.

.EXAMPLE
    Get-ChildItem "C:\Documents\*.pdf" | Add-DocumentToQueue -EnrichCode -EnrichFormula -EnrichPictureDescription
    Adds all PDF files with comprehensive enrichment features enabled.

.NOTES
    Supported file formats: .pdf, .docx, .xlsx, .pptx, .md, .html, .xhtml, .csv, .png, .jpg, .jpeg, .tiff, .tif, .bmp, .webp
#>
function Add-DocumentToQueue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]$Path,
        [string]$ExportFormat = 'markdown',
        [switch]$EmbedImages,
        [switch]$EnrichCode,
        [switch]$EnrichFormula,
        [switch]$EnrichPictureClasses,
        [switch]$EnrichPictureDescription
    )

    process {
        foreach ($filePath in $Path) {
            if (Test-Path $filePath) {
                $fileInfo = Get-Item $filePath
                $supportedFormats = @('.pdf', '.docx', '.xlsx', '.pptx', '.md', '.html', '.xhtml', '.csv', '.png', '.jpg', '.jpeg', '.tiff', '.tif', '.bmp', '.webp')

                if ($fileInfo.Extension -notin $supportedFormats) {
                    Write-Warning "Unsupported format: $($fileInfo.Extension)"
                    continue
                }

                $item = @{
                    Id                       = [guid]::NewGuid().ToString()
                    FilePath                 = $fileInfo.FullName
                    FileName                 = $fileInfo.Name
                    ExportFormat             = $ExportFormat
                    EmbedImages              = $EmbedImages.IsPresent
                    EnrichCode               = $EnrichCode.IsPresent
                    EnrichFormula            = $EnrichFormula.IsPresent
                    EnrichPictureClasses     = $EnrichPictureClasses.IsPresent
                    EnrichPictureDescription = $EnrichPictureDescription.IsPresent
                    Status                   = 'Ready'
                    UploadedTime             = Get-Date
                }

                # Don't add to processing queue yet - just store status
                Update-ItemStatus $item.Id $item

                Write-Host "Queued: $($fileInfo.Name) (ID: $($item.Id))" -ForegroundColor Green
                Write-Output $item.Id
            }
        }
    }
}

function Start-DocumentConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DocumentId,
        [string]$ExportFormat,
        [switch]$EmbedImages,
        [switch]$EnrichCode,
        [switch]$EnrichFormula,
        [switch]$EnrichPictureClasses,
        [switch]$EnrichPictureDescription
    )

    $allStatus = Get-ProcessingStatus
    $documentStatus = $allStatus[$DocumentId]

    if (-not $documentStatus) {
        Write-Error "Document not found: $DocumentId"
        return $false
    }

    if ($documentStatus.Status -ne 'Ready') {
        Write-Warning "Document $DocumentId is not in Ready status (current: $($documentStatus.Status))"
        return $false
    }

    # Update export format if provided
    if ($ExportFormat) {
        $documentStatus.ExportFormat = $ExportFormat
    }

    # Create queue item for processing
    $queueItem = @{
        Id                       = $DocumentId
        FilePath                 = $documentStatus.FilePath
        FileName                 = $documentStatus.FileName
        ExportFormat             = $documentStatus.ExportFormat
        EmbedImages              = $EmbedImages.IsPresent
        EnrichCode               = $EnrichCode.IsPresent
        EnrichFormula            = $EnrichFormula.IsPresent
        EnrichPictureClasses     = $EnrichPictureClasses.IsPresent
        EnrichPictureDescription = $EnrichPictureDescription.IsPresent
        Status                   = 'Queued'
        QueuedTime               = Get-Date
        UploadedTime             = $documentStatus.UploadedTime
    }

    # Add to processing queue and update status
    Add-QueueItem $queueItem
    Update-ItemStatus $DocumentId @{
        Status                   = 'Queued'
        QueuedTime               = Get-Date
        ExportFormat             = $documentStatus.ExportFormat
        EmbedImages              = $EmbedImages.IsPresent
        EnrichCode               = $EnrichCode.IsPresent
        EnrichFormula            = $EnrichFormula.IsPresent
        EnrichPictureClasses     = $EnrichPictureClasses.IsPresent
        EnrichPictureDescription = $EnrichPictureDescription.IsPresent
    }

    Write-Host "Started conversion for: $($documentStatus.FileName) (ID: $DocumentId)" -ForegroundColor Green
    return $true
}

<#
.SYNOPSIS
    Starts the background document processor service.

.DESCRIPTION
    Launches the document processing service that continuously monitors the queue for new documents and processes them using the Python Docling library. This function runs indefinitely in a loop, processing queued documents one at a time.

.EXAMPLE
    Start-DocumentProcessor

    Starts the document processor in the current PowerShell session.

.EXAMPLE
    Start-Job -ScriptBlock { Import-Module PSDocling; Start-DocumentProcessor }

    Starts the document processor in a background job.

.NOTES
    This function should typically be run in a separate PowerShell process or background job.
    The processor monitors the queue file at $env:TEMP\docling_queue.json and updates
    status information in $env:TEMP\docling_status.json.

    Processing includes:
    - File validation and format detection
    - Progress tracking with estimated completion times
    - Python Docling integration for document conversion
    - Output file generation in the ProcessedDocuments directory
    - Error handling and status reporting

.LINK
    Add-DocumentToQueue
    Get-DoclingSystemStatus
#>
function Start-DocumentProcessor {
    [CmdletBinding()]
    param()

    Write-Host "Document processor started" -ForegroundColor Green

    while ($true) {
        $item = Get-NextQueueItem
        if ($item) {
            Write-Host "Processing: $($item.FileName)" -ForegroundColor Yellow

            # Get file size for progress estimation
            $fileSize = (Get-Item $item.FilePath).Length
            $estimatedDurationMs = [Math]::Max(30000, [Math]::Min(300000, $fileSize / 1024 * 1000)) # 30s to 5min based on file size

            # Update status with progress tracking
            Update-ItemStatus $item.Id @{
                Status            = 'Processing'
                StartTime         = Get-Date
                Progress          = 0
                FileSize          = $fileSize
                EstimatedDuration = $estimatedDurationMs
            }

            # Create output directory
            $outputDir = Join-Path $script:DoclingSystem.OutputDirectory $item.Id
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($item.FileName)

            # Determine file extension based on export format
            $extension = switch ($item.ExportFormat) {
                'markdown' { '.md' }
                'html' { '.html' }
                'json' { '.json' }
                'text' { '.txt' }
                'doctags' { '.xml' }
                default { '.md' }
            }

            $outputFile = Join-Path $outputDir "$baseName$extension"

            $success = $false

            # Initialize all variables to prevent $null reference issues
            $stdout = $null
            $stderr = $null
            $pythonSuccess = $false
            $outputExists = $false
            $imagesExtracted = 0
            $imagesDirectory = $null
            $processCompletedNormally = $false
            $processExitCode = -1
            $processTerminatedEarly = $false

            try {
                if ($script:DoclingSystem.PythonAvailable) {
                    # Create Python conversion script
                    $pyScript = @"
import sys
import json
import re
import os
import base64
from pathlib import Path
from urllib.parse import quote
from datetime import datetime

try:
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.datamodel.base_models import InputFormat
    from docling_core.types.doc import ImageRefMode

    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    export_format = sys.argv[3] if len(sys.argv) > 3 else 'markdown'
    embed_images = sys.argv[4].lower() == 'true' if len(sys.argv) > 4 else False
    enrich_code = sys.argv[5].lower() == 'true' if len(sys.argv) > 5 else False
    enrich_formula = sys.argv[6].lower() == 'true' if len(sys.argv) > 6 else False
    enrich_picture_classes = sys.argv[7].lower() == 'true' if len(sys.argv) > 7 else False
    enrich_picture_description = sys.argv[8].lower() == 'true' if len(sys.argv) > 8 else False

    # Configure Docling for proper image extraction and enrichments
    pipeline_options = PdfPipelineOptions()
    pipeline_options.images_scale = 2.0  # Higher resolution images
    pipeline_options.generate_page_images = True
    pipeline_options.generate_picture_images = True  # Enable image extraction!

    # Configure enrichments
    pipeline_options.do_code_enrichment = enrich_code
    pipeline_options.do_formula_enrichment = enrich_formula
    pipeline_options.do_picture_classification = enrich_picture_classes
    pipeline_options.do_picture_description = enrich_picture_description

    # Configure Granite Vision model for picture description if enabled
    if enrich_picture_description:
        try:
            from docling.datamodel.pipeline_options import granite_picture_description
            pipeline_options.picture_description_options = granite_picture_description
        except ImportError:
            print("Warning: Granite Vision model not available, picture description disabled", file=sys.stderr)

    print("Creating DocumentConverter...", file=sys.stderr)
    converter = DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
        }
    )
    print(f"Starting conversion of {src}...", file=sys.stderr)
    result = converter.convert(str(src))
    print("Conversion completed", file=sys.stderr)

    # Create images directory next to output file
    images_dir = dst.parent / f"{dst.stem}_images"
    images_dir.mkdir(exist_ok=True)

    # Helper function to save images and update references
    def save_images_and_update_content(content, base_format='markdown'):
        images_extracted = 0
        vector_graphics_found = 0
        updated_content = content
        image_replacements = []

        try:
            # Check for pictures in the document
            if hasattr(result.document, 'pictures') and result.document.pictures:
                print(f"Found {len(result.document.pictures)} pictures in document", file=sys.stderr)

                for i, picture in enumerate(result.document.pictures):
                    try:
                        # Try to get extractable raster image
                        pil_image = picture.get_image(result.document)

                        if pil_image:
                            # We have a raster image - extract it
                            image_filename = f"image_{images_extracted + 1:03d}.png"
                            image_path = images_dir / image_filename
                            pil_image.save(str(image_path), 'PNG')

                            # Create reference in content
                            relative_path = f"{dst.stem}_images/{image_filename}"
                            if base_format == 'markdown':
                                img_reference = f"![Image {images_extracted + 1}]({relative_path})"
                            elif base_format == 'html':
                                img_reference = f'<img src="{relative_path}" alt="Image {images_extracted + 1}" />'
                            else:
                                img_reference = f"[Image: {relative_path}]"

                            # Store image reference for placeholder replacement
                            image_replacements.append(img_reference)
                            print(f"Successfully extracted raster image {images_extracted + 1} to {image_path}", file=sys.stderr)
                            images_extracted += 1
                        else:
                            # This is a vector graphic or non-extractable image element
                            vector_graphics_found += 1

                            # Get position information if available
                            position_info = ""
                            if hasattr(picture, 'prov') and picture.prov:
                                prov = picture.prov[0]
                                page = prov.page_no
                                bbox = prov.bbox
                                position_info = f" (Page {page}, position {bbox.l:.0f},{bbox.t:.0f}-{bbox.r:.0f},{bbox.b:.0f})"

                            # Create informative placeholder
                            if base_format == 'markdown':
                                placeholder = f"![Vector Graphic {vector_graphics_found}](# \"Vector graphic or logo detected{position_info}\")"
                            elif base_format == 'html':
                                placeholder = f'<!-- Vector Graphic {vector_graphics_found}: Non-extractable image element{position_info} -->'
                            else:
                                placeholder = f"[Vector Graphic {vector_graphics_found}: Non-extractable image element{position_info}]"

                            # Store placeholder for replacement
                            image_replacements.append(placeholder)
                            print(f"Found vector graphic {vector_graphics_found}{position_info}", file=sys.stderr)

                    except Exception as img_error:
                        print(f"Warning: Could not process picture {i + 1}: {img_error}", file=sys.stderr)

            # Report summary
            if images_extracted > 0:
                print(f"Total raster images extracted: {images_extracted}", file=sys.stderr)
            if vector_graphics_found > 0:
                print(f"Total vector graphics detected: {vector_graphics_found}", file=sys.stderr)
            if images_extracted == 0 and vector_graphics_found == 0:
                print("No images or graphics found in document", file=sys.stderr)

        except Exception as e:
            print(f"Warning: Image processing failed: {e}", file=sys.stderr)

        # Replace <!-- image --> placeholders with actual image references
        if image_replacements:
            import re
            replacement_index = 0
            def replace_image_placeholder(match):
                nonlocal replacement_index
                if replacement_index < len(image_replacements):
                    replacement = image_replacements[replacement_index]
                    replacement_index += 1
                    return replacement
                else:
                    return "<!-- No image data available -->"

            # Perform the replacement
            updated_content = re.sub(r'<!-- image -->', replace_image_placeholder, updated_content, flags=re.IGNORECASE)
            print(f"Replaced {replacement_index} image placeholders", file=sys.stderr)

        # Return both the content and count of actual extracted images
        return updated_content, images_extracted

    # Use Docling's native image handling based on embed_images setting
    image_mode = ImageRefMode.EMBEDDED if embed_images else ImageRefMode.REFERENCED
    images_extracted = len(result.document.pictures) if hasattr(result.document, 'pictures') else 0

    dst.parent.mkdir(parents=True, exist_ok=True)

    # Export using our custom methods to control directory structure and filenames
    if export_format == 'markdown':
        if image_mode == ImageRefMode.EMBEDDED:
            content = result.document.export_to_markdown(image_mode=ImageRefMode.EMBEDDED)
        else:
            content = result.document.export_to_markdown()
            content, images_extracted = save_images_and_update_content(content, 'markdown')
        dst.write_text(content, encoding='utf-8')
        print(f"Saved markdown with custom image handling", file=sys.stderr)
    elif export_format == 'html':
        if image_mode == ImageRefMode.EMBEDDED:
            content = result.document.export_to_html(image_mode=ImageRefMode.EMBEDDED)
        else:
            content = result.document.export_to_html()
            content, images_extracted = save_images_and_update_content(content, 'html')
        dst.write_text(content, encoding='utf-8')
        print(f"Saved HTML with custom image handling", file=sys.stderr)
    elif export_format == 'json':
        import json
        doc_dict = result.document.export_to_dict()
        content = json.dumps(doc_dict, indent=2, ensure_ascii=False)
        dst.write_text(content, encoding='utf-8')
        print(f"Saved JSON (images in document structure)", file=sys.stderr)
    elif export_format == 'text':
        # For text format, use markdown export and convert
        try:
            content = result.document.export_to_text()
        except AttributeError:
            # Fallback: get markdown and convert to text
            md_content = result.document.export_to_markdown(image_mode=ImageRefMode.EMBEDDED)

            # Simple markdown to text conversion
            import re
            content = re.sub(r'[#*`_]', '', md_content)
            content = re.sub(r'!\[(.*?)\]\(data:.*?\)', r'[Embedded Image: \1]', content)
            content = re.sub(r'!\[(.*?)\]\((.*?)\)', r'[Image: \1 - \2]', content)
            content = re.sub(r'\[.*?\]\((.*?)\)', r'[Link: \1]', content)
            content = re.sub(r'\n+', '\n', content).strip()

        dst.write_text(content, encoding='utf-8')
        print(f"Saved text format", file=sys.stderr)
    elif export_format == 'doctags':
        import xml.etree.ElementTree as ET
        from xml.sax.saxutils import escape
        import json

        try:
            # Try the native export_to_doctags first
            raw_content = result.document.export_to_doctags()
            print(f"Raw DocTags content length: {len(raw_content)}", file=sys.stderr)

            # Always wrap DocTags in proper XML structure due to known malformed output
            # The Docling export_to_doctags produces unclosed tags which break XML parsers
            escaped_content = escape(raw_content)
            content = f'''<?xml version="1.0" encoding="UTF-8"?>
<document>
    <metadata>
        <source>Docling DocTags Export</source>
        <generated>{datetime.now().isoformat()}</generated>
    </metadata>
    <doctags><![CDATA[{escaped_content}]]></doctags>
</document>'''
            print(f"Wrapped DocTags in valid XML structure", file=sys.stderr)
        except (AttributeError, Exception) as e:
            print(f"DocTags export failed: {e}, using fallback", file=sys.stderr)
            # Fallback: create a proper XML representation from document structure
            doc_dict = result.document.export_to_dict()

            content = f'''<?xml version="1.0" encoding="UTF-8"?>
<document>
    <metadata>
        <title>{escape(str(doc_dict.get('name', 'Unknown Document')))}</title>
        <source>Docling JSON Fallback</source>
        <generated>{datetime.now().isoformat()}</generated>
    </metadata>
    <content><![CDATA[{json.dumps(doc_dict, indent=2)}]]></content>
</document>'''

        dst.write_text(content, encoding='utf-8')
        print(f"Saved doctags format", file=sys.stderr)
    else:
        raise ValueError(f'Unsupported export format: {export_format}')

    # Report image handling
    if embed_images:
        print(f"Images embedded directly in {export_format} file", file=sys.stderr)
    else:
        print(f"Images saved as separate files with references", file=sys.stderr)

    print(json.dumps({
        'success': True,
        'format': export_format,
        'output_file': str(dst),
        'images_extracted': images_extracted,
        'images_directory': str(images_dir) if images_extracted > 0 else None
    }))

except Exception as e:
    print(json.dumps({'success': False, 'error': str(e)}))
    sys.exit(1)
"@

                    $tempPy = Join-Path $env:TEMP "docling_$([guid]::NewGuid().ToString('N')[0..7] -join '').py"
                    $pyScript | Set-Content $tempPy -Encoding UTF8

                    try {
                        # Start Python process with timeout (10 minutes max)
                        # Use single argument string to properly handle spaces in filenames
                        $exportFormat = if ($item.ExportFormat) { $item.ExportFormat } else { 'markdown' }
                        $embedImages = if ($item.EmbedImages) { 'true' } else { 'false' }
                        $enrichCode = if ($item.EnrichCode) { 'true' } else { 'false' }
                        $enrichFormula = if ($item.EnrichFormula) { 'true' } else { 'false' }
                        $enrichPictureClasses = if ($item.EnrichPictureClasses) { 'true' } else { 'false' }
                        $enrichPictureDescription = if ($item.EnrichPictureDescription) { 'true' } else { 'false' }
                        $arguments = "`"$tempPy`" `"$($item.FilePath)`" `"$outputFile`" `"$exportFormat`" `"$embedImages`" `"$enrichCode`" `"$enrichFormula`" `"$enrichPictureClasses`" `"$enrichPictureDescription`""
                        $process = Start-Process python -ArgumentList $arguments -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\docling_output.txt" -RedirectStandardError "$env:TEMP\docling_error.txt"

                        # Monitor process with progress updates
                        $startTime = Get-Date
                        $finished = $false
                        $lastProgressUpdate = 0

                        while (-not $process.HasExited) {
                            Start-Sleep -Milliseconds 1000  # Check every second
                            $elapsed = (Get-Date) - $startTime
                            $elapsedMs = $elapsed.TotalMilliseconds

                            # Check for timeout - extend for AI model enrichments (Granite Vision, CodeFormula model loading)
                            $timeoutSeconds = if ($item.EnrichPictureDescription) {
                                2400  # 40 min for Granite Vision
                            } elseif ($item.EnrichCode -or $item.EnrichFormula) {
                                1800  # 30 min for Code/Formula Understanding
                            } else {
                                1200  # 20 min for standard processing
                            }
                            if ($elapsed.TotalSeconds -gt $timeoutSeconds) {
                                $timeoutMinutes = $timeoutSeconds / 60
                                Write-Host "Process timeout for $($item.FileName) after $timeoutMinutes minutes, terminating..." -ForegroundColor Yellow
                                $processTerminatedEarly = $true
                                try {
                                    $process.Kill()
                                }
                                catch {
                                    Write-Host "Could not kill process: $($_.Exception.Message)" -ForegroundColor Red
                                }
                                break
                            }

                            # Check for early Python completion/failure by examining output files
                            $outputFileExists = Test-Path "$env:TEMP\docling_output.txt" -ErrorAction SilentlyContinue
                            if ($outputFileExists) {
                                $stdout = Get-Content "$env:TEMP\docling_output.txt" -Raw -ErrorAction SilentlyContinue
                                if ($stdout) {
                                    try {
                                        $jsonResult = $stdout | ConvertFrom-Json
                                        if ($jsonResult.success -eq $true -or $jsonResult.success -eq $false) {
                                            # Python has finished (either success or failure) - break out of monitoring
                                            Write-Host "Python process completed, breaking monitoring loop..." -ForegroundColor Green
                                            break
                                        }
                                    }
                                    catch {
                                        # Ignore JSON parsing errors, continue monitoring
                                    }
                                }
                            }

                            # Calculate and update progress with guards - special handling for AI model enrichments
                            if ($item.EnrichPictureDescription) {
                                # Picture Description (Granite Vision) - extended timeline with model loading phases
                                if ($elapsed.TotalSeconds -lt 300) {
                                    # First 5 minutes: Model download/loading - slow progress to 25%
                                    $progress = ($elapsed.TotalSeconds / 300.0) * 25.0
                                } elseif ($elapsed.TotalSeconds -lt 900) {
                                    # Next 10 minutes: Model processing - 25% to 80%
                                    $progress = 25.0 + (($elapsed.TotalSeconds - 300.0) / 600.0) * 55.0
                                } else {
                                    # Final phase: slow progress to 95%
                                    $progress = 80.0 + [Math]::Min(15.0, (($elapsed.TotalSeconds - 900.0) / 600.0) * 15.0)
                                }
                            } elseif ($item.EnrichCode -or $item.EnrichFormula) {
                                # Code/Formula Understanding (CodeFormulaV2) - similar timeline to Granite Vision
                                if ($elapsed.TotalSeconds -lt 180) {
                                    # First 3 minutes: Model download/loading - slow progress to 20%
                                    $progress = ($elapsed.TotalSeconds / 180.0) * 20.0
                                } elseif ($elapsed.TotalSeconds -lt 600) {
                                    # Next 7 minutes: Model processing - 20% to 85%
                                    $progress = 20.0 + (($elapsed.TotalSeconds - 180.0) / 420.0) * 65.0
                                } else {
                                    # Final phase: slow progress to 95%
                                    $progress = 85.0 + [Math]::Min(10.0, (($elapsed.TotalSeconds - 600.0) / 300.0) * 10.0)
                                }
                            } elseif ($estimatedDurationMs -gt 0) {
                                $progress = [Math]::Min(95.0, ([double]($elapsedMs) / [double]($estimatedDurationMs)) * 100.0)
                            }
                            else {
                                $progress = [Math]::Min(95.0, ([double]($elapsedMs) / 60000.0) * 100.0)
                            }

                            # Only update if progress changed significantly
                            if ([Math]::Abs($progress - $lastProgressUpdate) -gt 1.0) {
                                Update-ItemStatus $item.Id @{
                                    Progress    = [Math]::Round($progress, 1)
                                    ElapsedTime = $elapsedMs
                                    LastUpdate  = Get-Date
                                }
                                $lastProgressUpdate = $progress
                            }
                        }

                        # Process has exited - wait for file writes to complete
                        Start-Sleep -Milliseconds 500

                        $finished = $true

                        if ($finished) {
                            # Process has exited - now check results AFTER process completion
                            # Wait for process to fully exit before accessing ExitCode
                            if (-not $process.HasExited) {
                                try {
                                    # Wait up to 5 seconds for process to fully exit
                                    $process.WaitForExit(5000) | Out-Null
                                }
                                catch {
                                    Write-Warning "Failed to wait for process exit: $($_.Exception.Message)"
                                }
                            }

                            # Now safely access ExitCode
                            if ($process.HasExited) {
                                $processExitCode = $process.ExitCode
                                $processCompletedNormally = $true
                            }
                            else {
                                $processExitCode = -1
                                $processCompletedNormally = $false
                            }

                            # Read output files after process completion
                            $stdout = Get-Content "$env:TEMP\docling_output.txt" -Raw -ErrorAction SilentlyContinue
                            $stderr = Get-Content "$env:TEMP\docling_error.txt" -Raw -ErrorAction SilentlyContinue

                            # Check for success in Python output
                            $pythonSuccess = $false
                            if ($stdout) {
                                try {
                                    $jsonResult = $stdout | ConvertFrom-Json
                                    $pythonSuccess = $jsonResult.success -eq $true
                                    if ($jsonResult.images_extracted) {
                                        $imagesExtracted = $jsonResult.images_extracted
                                    }
                                    if ($jsonResult.images_directory) {
                                        $imagesDirectory = $jsonResult.images_directory
                                    }
                                }
                                catch {
                                    # Fallback to regex check
                                    $pythonSuccess = $stdout -match '"success".*true'
                                }
                            }

                            # Check if output file exists and has content
                            $outputExists = (Test-Path $outputFile -ErrorAction SilentlyContinue) -and
                            ((Get-Item $outputFile -ErrorAction SilentlyContinue).Length -gt 0)

                            # Determine success based on Python results and output
                            if ($pythonSuccess -and $outputExists) {
                                $success = $true
                            }
                            elseif ($processCompletedNormally -and ($processExitCode -eq 0) -and $outputExists) {
                                $success = $true
                            }
                            else {
                                $success = $false
                                $errorMsg = "Document processing failed"
                                if (-not $processCompletedNormally) {
                                    $errorMsg += " - Process did not complete normally"
                                }
                                if ($processExitCode -ne 0) {
                                    $errorMsg += " - Exit code: $processExitCode"
                                }
                                if ($stderr) {
                                    $errorMsg += " - Python error: $($stderr.Substring(0, [Math]::Min(500, $stderr.Length)))"
                                }
                                throw $errorMsg
                            }
                        }

                        # Clean up temp files
                        Remove-Item "$env:TEMP\docling_output.txt" -Force -ErrorAction SilentlyContinue
                        Remove-Item "$env:TEMP\docling_error.txt" -Force -ErrorAction SilentlyContinue

                    }
                    finally {
                        Remove-Item $tempPy -Force -ErrorAction SilentlyContinue
                    }
                }
                else {
                    # Simulation mode - properly initialize variables
                    Start-Sleep 2
                    "Simulated conversion of: $($item.FileName)`nGenerated at: $(Get-Date)" | Set-Content $outputFile -Encoding UTF8
                    $success = $true
                    $processCompletedNormally = $true
                    $processExitCode = 0
                    $pythonSuccess = $true
                    $outputExists = Test-Path $outputFile -ErrorAction SilentlyContinue
                    $imagesExtracted = 0
                    $imagesDirectory = $null
                }

                if ($success) {
                    $statusUpdate = @{
                        Status     = 'Completed'
                        OutputFile = $outputFile
                        EndTime    = Get-Date
                        Progress   = 100
                    }

                    # Add image extraction info if available
                    if ($imagesExtracted -gt 0) {
                        $statusUpdate.ImagesExtracted = $imagesExtracted
                        $statusUpdate.ImagesDirectory = $imagesDirectory
                        Write-Host "Completed: $($item.FileName) ($imagesExtracted images extracted)" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Completed: $($item.FileName)" -ForegroundColor Green
                    }

                    Update-ItemStatus $item.Id $statusUpdate
                }
                else {
                    # Provide more detailed error information
                    $errorMsg = "Conversion failed. "
                    if (-not $outputExists) {
                        $errorMsg += "Output file not created or empty. "
                    }
                    if (-not $pythonSuccess) {
                        $errorMsg += "Python script did not report success. "
                    }
                    if ($stderr) {
                        $errorMsg += "Python stderr: $stderr"
                    }
                    if ($stdout) {
                        $errorMsg += "Python stdout: $stdout"
                    }
                    throw $errorMsg
                }

            }
            catch {
                # Capture detailed error information
                $errorDetails = @{
                    ExceptionType    = $_.Exception.GetType().Name
                    StackTrace       = $_.Exception.StackTrace
                    InnerException   = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $null }
                    ScriptStackTrace = $_.ScriptStackTrace
                }

                # Try to get stderr if it exists
                $stderr = ""
                try {
                    if (Test-Path "$env:TEMP\docling_error.txt") {
                        $stderr = Get-Content "$env:TEMP\docling_error.txt" -Raw -ErrorAction SilentlyContinue
                    }
                }
                catch { }

                Update-ItemStatus $item.Id @{
                    Status       = 'Error'
                    Error        = $_.Exception.Message
                    ErrorDetails = $errorDetails
                    StdErr       = $stderr
                    EndTime      = Get-Date
                    Progress     = 0
                }
                Write-Host "Error processing $($item.FileName): $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        Start-Sleep 2
    }
}

<#
.SYNOPSIS
    Starts the REST API server for the PSDocling system.

.DESCRIPTION
    Launches an HTTP listener that provides REST API endpoints for file upload, status monitoring, document management, and health checks. The server handles CORS and provides JSON responses for all endpoints.

.PARAMETER Port
    The port number to bind the HTTP listener to. Default is 8080.

.EXAMPLE
    Start-APIServer
    Starts the API server on the default port 8080.

.EXAMPLE
    Start-APIServer -Port 9080
    Starts the API server on port 9080.

.NOTES
    Available API endpoints:
    - POST /api/upload - Upload documents for processing
    - GET /api/status - Get system status
    - GET /api/documents - List processed documents
    - GET /api/files/* - Download processed files
    - GET /api/health - Health check endpoint

    The server runs indefinitely until stopped with Ctrl+C or process termination.
#>
function Start-APIServer {
    [CmdletBinding()]
    param([int]$Port = 8080)

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$Port/")

    try {
        $listener.Start()
        Write-Host "API Server started on port $Port" -ForegroundColor Green

        while ($listener.IsListening) {
            try {
                $context = $listener.GetContext()
                $request = $context.Request
                $response = $context.Response

                # CORS
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")

                if ($request.HttpMethod -eq "OPTIONS") {
                    $response.StatusCode = 200
                    $response.Close()
                    continue
                }

                $responseContent = ""
                $path = $request.Url.LocalPath

                switch -Regex ($path) {
                    '^/api/health$' {
                        $responseContent = @{
                            status      = 'healthy'
                            timestamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                            version     = $script:DoclingSystem.Version
                            codeVersion = "2025-09-21-updated"
                        } | ConvertTo-Json
                    }

                    '^/api/status$' {
                        $queue = Get-QueueItems
                        $allStatus = Get-ProcessingStatus
                        $queued = $allStatus.Values | Where-Object { $_.Status -eq 'Queued' }
                        $processing = $allStatus.Values | Where-Object { $_.Status -eq 'Processing' }
                        $completed = $allStatus.Values | Where-Object { $_.Status -eq 'Completed' }
                        $errors = $allStatus.Values | Where-Object { $_.Status -eq 'Error' }

                        $responseContent = @{
                            QueuedCount     = @($queued).Count
                            ProcessingCount = @($processing).Count
                            CompletedCount  = @($completed).Count
                            ErrorCount      = @($errors).Count
                            TotalItems      = $allStatus.Count
                        } | ConvertTo-Json
                    }

                    '^/api/documents$' {
                        $allStatus = Get-ProcessingStatus
                        $documents = @()

                        if ($allStatus -and $allStatus.Count -gt 0) {
                            $allStatus.GetEnumerator() | ForEach-Object {
                                $doc = @{
                                    id           = $_.Key
                                    fileName     = if ($_.Value.FileName) { $_.Value.FileName } else { "unknown" }
                                    status       = $_.Value.Status
                                    exportFormat = if ($_.Value.ExportFormat) { $_.Value.ExportFormat } else { "markdown" }
                                }

                                # Add enrichment options - always include with default false if not set
                                $doc.embedImages = if ($_.Value.EmbedImages) { [bool]$_.Value.EmbedImages } else { $false }
                                $doc.enrichCode = if ($_.Value.EnrichCode) { [bool]$_.Value.EnrichCode } else { $false }
                                $doc.enrichFormula = if ($_.Value.EnrichFormula) { [bool]$_.Value.EnrichFormula } else { $false }
                                $doc.enrichPictureClasses = if ($_.Value.EnrichPictureClasses) { [bool]$_.Value.EnrichPictureClasses } else { $false }
                                $doc.enrichPictureDescription = if ($_.Value.EnrichPictureDescription) { [bool]$_.Value.EnrichPictureDescription } else { $false }

                                # Add progress data if available
                                if ($_.Value.Progress) { $doc.progress = $_.Value.Progress }
                                if ($_.Value.EstimatedDuration) { $doc.estimatedDuration = $_.Value.EstimatedDuration }
                                if ($_.Value.ElapsedTime) { $doc.elapsedTime = $_.Value.ElapsedTime }
                                if ($_.Value.FileSize) { $doc.fileSize = $_.Value.FileSize }
                                if ($_.Value.StartTime) { $doc.startTime = $_.Value.StartTime }
                                if ($_.Value.LastUpdate) { $doc.lastUpdate = $_.Value.LastUpdate }

                                $documents += $doc
                            }
                        }

                        $responseContent = if ($documents.Count -eq 0) { "[]" } else {
                            # Ensure array format even for single item
                            if ($documents.Count -eq 1) {
                                "[$($documents | ConvertTo-Json -Depth 10)]"
                            }
                            else {
                                $documents | ConvertTo-Json -Depth 10
                            }
                        }
                    }

                    '^/api/files$' {
                        # Use ArrayList to ensure array behavior
                        $files = New-Object System.Collections.ArrayList
                        $processedDir = $script:DoclingSystem.OutputDirectory

                        try {
                            if (Test-Path $processedDir) {
                                Get-ChildItem $processedDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                                    $docId = $_.Name
                                    try {
                                        Get-ChildItem $_.FullName -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml') } | ForEach-Object {
                                            $filePath = $_.FullName
                                            $fileName = $_.Name
                                            $fileSize = [math]::Round($_.Length / 1KB, 2)
                                            $lastModified = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")

                                            $fileObj = @{
                                                id           = $docId
                                                fileName     = $fileName
                                                filePath     = $filePath
                                                fileSize     = "$fileSize KB"
                                                lastModified = $lastModified
                                                downloadUrl  = "/api/file/" + $docId + "?file=" + $fileName
                                            }
                                            [void]$files.Add($fileObj)
                                        }
                                    }
                                    catch {
                                        # Skip directories that can't be accessed
                                        Write-Warning "Could not access directory: $($_.FullName)"
                                    }
                                }
                            }

                            # Simple response - frontend handles array conversion
                            $responseContent = if ($files.Count -eq 0) {
                                "[]"
                            }
                            else {
                                $files | ConvertTo-Json -Depth 10
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to load processed files"
                                details = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }

                    '^/api/error/(.+)$' {
                        $id = $Matches[1]
                        $allStatus = Get-ProcessingStatus
                        $status = $allStatus[$id]

                        try {
                            # Debug information
                            $debugInfo = @{
                                requestedId   = $id
                                statusFound   = $status -ne $null
                                statusData    = $status
                                allStatusKeys = $allStatus.Keys -join ', '
                                statusType    = if ($status) { $status.GetType().Name } else { 'null' }
                            }

                            if (-not $status) {
                                $response.StatusCode = 404
                                $responseContent = @{
                                    error = "Document not found"
                                    debug = $debugInfo
                                } | ConvertTo-Json -Depth 10
                            }
                            elseif ($status.Status -eq 'Error') {
                                $errorDetails = @{
                                    id           = $id
                                    fileName     = $status.FileName
                                    status       = $status.Status
                                    error        = $status.Error
                                    errorDetails = $status.ErrorDetails
                                    stderr       = $status.StdErr
                                    queuedTime   = $status.QueuedTime
                                    startTime    = $status.StartTime
                                    endTime      = $status.EndTime
                                    debug        = $debugInfo
                                }
                                $responseContent = $errorDetails | ConvertTo-Json -Depth 10
                            }
                            else {
                                # For non-error states, provide helpful status information instead of an error
                                $currentInfo = @{
                                    id                = $id
                                    fileName          = $status.FileName
                                    currentStatus     = $status.Status
                                    message           = switch ($status.Status) {
                                        'Processing' { "Document is currently being processed. Progress: $($status.Progress)%" }
                                        'Queued' { "Document is waiting in the processing queue" }
                                        'Ready' { "Document is ready for processing" }
                                        'Completed' { "Document processing completed successfully" }
                                        default { "Document is in $($status.Status) state" }
                                    }
                                    queuedTime        = $status.QueuedTime
                                    startTime         = $status.StartTime
                                    progress          = $status.Progress
                                    estimatedDuration = $status.EstimatedDuration
                                    elapsedTime       = $status.ElapsedTime
                                }

                                # Return 200 OK with status information instead of 400 error
                                $responseContent = $currentInfo | ConvertTo-Json -Depth 10
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to retrieve error details"
                                details = $_.Exception.Message
                                debug   = $debugInfo
                            } | ConvertTo-Json -Depth 10
                        }
                    }

                    '^/api/file/(.+)$' {
                        $docId = $Matches[1]
                        $processedDir = $script:DoclingSystem.OutputDirectory
                        $docDir = Join-Path $processedDir $docId

                        try {
                            # Check if a specific file is requested via query parameter
                            $requestedFile = $null
                            if ($request.Url.Query) {
                                $queryParams = [System.Web.HttpUtility]::ParseQueryString($request.Url.Query)
                                $requestedFile = $queryParams['file']
                            }

                            if (Test-Path $docDir) {
                                if ($requestedFile) {
                                    # Serve specific file with security validation
                                    try {
                                        $secureFileName = Get-SecureFileName -FileName $requestedFile -AllowedExtensions @('.md', '.html', '.json', '.txt', '.xml')
                                        $filePath = Join-Path $docDir $secureFileName

                                        # Additional security check: ensure the resolved path is still within the document directory
                                        $resolvedPath = [System.IO.Path]::GetFullPath($filePath)
                                        $resolvedDocDir = [System.IO.Path]::GetFullPath($docDir)

                                        if (-not $resolvedPath.StartsWith($resolvedDocDir + [System.IO.Path]::DirectorySeparatorChar) -and $resolvedPath -ne $resolvedDocDir) {
                                            throw "Path traversal attempt detected"
                                        }

                                        if (Test-Path $filePath) {
                                            $outputFile = Get-Item $filePath
                                        }
                                        else {
                                            $outputFile = $null
                                        }
                                    }
                                    catch {
                                        Write-Warning "Security violation in file request: $($_.Exception.Message)"
                                        $outputFile = $null
                                    }
                                }
                                else {
                                    # Fallback: serve first available file (for Processing Results)
                                    $outputFile = Get-ChildItem $docDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml') } | Select-Object -First 1
                                }

                                if ($outputFile -and $outputFile.Extension -in @('.md', '.html', '.json', '.txt', '.xml')) {
                                    $bytes = [System.IO.File]::ReadAllBytes($outputFile.FullName)
                                    $contentType = switch ($outputFile.Extension) {
                                        '.md' { 'text/markdown; charset=utf-8' }
                                        '.html' { 'text/html; charset=utf-8' }
                                        '.json' { 'application/json; charset=utf-8' }
                                        '.txt' { 'text/plain; charset=utf-8' }
                                        '.xml' { 'application/xml; charset=utf-8' }
                                        default { 'text/plain; charset=utf-8' }
                                    }
                                    $response.ContentType = $contentType
                                    $response.ContentLength64 = $bytes.Length
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                    $response.OutputStream.Close()
                                    $response.Close()
                                    continue
                                }
                            }

                            $response.StatusCode = 404
                            $responseContent = @{ error = "File not found" } | ConvertTo-Json
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to serve file"
                                details = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }

                    '^/api/upload$' {
                        if ($request.HttpMethod -eq 'POST') {
                            try {
                                $reader = New-Object System.IO.StreamReader($request.InputStream)
                                $body = $reader.ReadToEnd()
                                $data = $body | ConvertFrom-Json

                                # Save file with security validation
                                try {
                                    $secureFileName = Get-SecureFileName -FileName $data.fileName
                                }
                                catch {
                                    $response.StatusCode = 400
                                    $responseContent = @{
                                        success = $false
                                        error   = "Invalid filename: $($_.Exception.Message)"
                                    } | ConvertTo-Json
                                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                    return
                                }

                                $uploadId = [guid]::NewGuid().ToString()
                                $uploadDir = Join-Path $script:DoclingSystem.TempDirectory $uploadId
                                New-Item -ItemType Directory -Force -Path $uploadDir | Out-Null

                                $filePath = Join-Path $uploadDir $secureFileName
                                [System.IO.File]::WriteAllBytes($filePath, [Convert]::FromBase64String($data.dataBase64))

                                # Prepare Add-DocumentToQueue parameters
                                $queueParams = @{
                                    Path = $filePath
                                }

                                # Add optional parameters if provided
                                if ($data.exportFormat) { $queueParams.ExportFormat = $data.exportFormat }
                                if ($data.embedImages -eq $true) { $queueParams.EmbedImages = $true }
                                if ($data.enrichCode -eq $true) { $queueParams.EnrichCode = $true }
                                if ($data.enrichFormula -eq $true) { $queueParams.EnrichFormula = $true }
                                if ($data.enrichPictureClasses -eq $true) { $queueParams.EnrichPictureClasses = $true }
                                if ($data.enrichPictureDescription -eq $true) { $queueParams.EnrichPictureDescription = $true }

                                # Queue with parameters
                                $queueId = Add-DocumentToQueue @queueParams

                                $responseContent = @{
                                    success    = $true
                                    documentId = $queueId
                                    message    = "Document uploaded and queued"
                                } | ConvertTo-Json

                            }
                            catch {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error   = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                    }

                    '^/api/reprocess$' {
                        if ($request.HttpMethod -eq 'POST') {
                            try {
                                $reader = New-Object System.IO.StreamReader($request.InputStream)
                                $body = $reader.ReadToEnd()
                                $data = $body | ConvertFrom-Json

                                $documentId = $data.documentId
                                $newFormat = $data.exportFormat
                                $embedImages = if ($data.embedImages) { $data.embedImages } else { $false }
                                $enrichCode = if ($data.enrichCode) { $data.enrichCode } else { $false }
                                $enrichFormula = if ($data.enrichFormula) { $data.enrichFormula } else { $false }
                                $enrichPictureClasses = if ($data.enrichPictureClasses) { $data.enrichPictureClasses } else { $false }
                                $enrichPictureDescription = if ($data.enrichPictureDescription) { $data.enrichPictureDescription } else { $false }

                                # Get the current document status
                                $allStatus = Get-ProcessingStatus
                                $currentStatus = $allStatus[$documentId]

                                if (-not $currentStatus) {
                                    $response.StatusCode = 404
                                    $responseContent = @{
                                        success = $false
                                        error   = "Document not found"
                                    } | ConvertTo-Json
                                }
                                else {
                                    # Create new queue item for reprocessing
                                    $reprocessItem = @{
                                        Id                       = $documentId  # Keep same ID to update existing entry
                                        FilePath                 = $currentStatus.FilePath
                                        FileName                 = $currentStatus.FileName
                                        ExportFormat             = $newFormat
                                        EmbedImages              = $embedImages
                                        EnrichCode               = $enrichCode
                                        EnrichFormula            = $enrichFormula
                                        EnrichPictureClasses     = $enrichPictureClasses
                                        EnrichPictureDescription = $enrichPictureDescription
                                        Status                   = 'Queued'
                                        QueuedTime               = Get-Date
                                        IsReprocess              = $true
                                    }

                                    # Add to queue and update status - preserve existing fields
                                    Add-QueueItem $reprocessItem
                                    Update-ItemStatus $documentId @{
                                        Status                   = 'Queued'
                                        ExportFormat             = $newFormat
                                        EmbedImages              = $embedImages
                                        EnrichCode               = $enrichCode
                                        EnrichFormula            = $enrichFormula
                                        EnrichPictureClasses     = $enrichPictureClasses
                                        EnrichPictureDescription = $enrichPictureDescription
                                        QueuedTime               = Get-Date
                                        IsReprocess              = $true
                                        # Preserve existing fields that might be needed
                                        FilePath                 = $currentStatus.FilePath
                                        FileName                 = $currentStatus.FileName
                                        OriginalQueuedTime       = $currentStatus.QueuedTime
                                        CompletedTime            = $currentStatus.EndTime
                                    }

                                    $responseContent = @{
                                        success    = $true
                                        documentId = $documentId
                                        message    = "Document queued for reprocessing with format: $newFormat"
                                    } | ConvertTo-Json
                                }
                            }
                            catch {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error   = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                    }

                    '^/api/start-conversion$' {
                        if ($request.HttpMethod -eq 'POST') {
                            try {
                                $reader = New-Object System.IO.StreamReader($request.InputStream)
                                $body = $reader.ReadToEnd()
                                $data = $body | ConvertFrom-Json

                                $documentId = $data.documentId
                                $exportFormat = $data.exportFormat
                                $embedImages = if ($data.embedImages) { $data.embedImages } else { $false }
                                $enrichCode = if ($data.enrichCode) { $data.enrichCode } else { $false }
                                $enrichFormula = if ($data.enrichFormula) { $data.enrichFormula } else { $false }
                                $enrichPictureClasses = if ($data.enrichPictureClasses) { $data.enrichPictureClasses } else { $false }
                                $enrichPictureDescription = if ($data.enrichPictureDescription) { $data.enrichPictureDescription } else { $false }

                                $success = Start-DocumentConversion -DocumentId $documentId -ExportFormat $exportFormat -EmbedImages:$embedImages -EnrichCode:$enrichCode -EnrichFormula:$enrichFormula -EnrichPictureClasses:$enrichPictureClasses -EnrichPictureDescription:$enrichPictureDescription

                                if ($success) {
                                    $responseContent = @{
                                        success    = $true
                                        documentId = $documentId
                                        message    = "Conversion started"
                                    } | ConvertTo-Json
                                }
                                else {
                                    $response.StatusCode = 400
                                    $responseContent = @{
                                        success = $false
                                        error   = "Failed to start conversion"
                                    } | ConvertTo-Json
                                }
                            }
                            catch {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error   = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                    }

                    '^/api/result/(.+)$' {
                        $id = $Matches[1]
                        $allStatus = Get-ProcessingStatus
                        $status = $allStatus[$id]

                        try {
                            if (-not $status) {
                                $response.StatusCode = 404
                                $responseContent = @{ error = "Document not found" } | ConvertTo-Json
                            }
                            elseif ($status.Status -eq 'Completed' -and $status.OutputFile -and (Test-Path $status.OutputFile)) {
                                $fileInfo = Get-Item $status.OutputFile
                                $fileExtension = [System.IO.Path]::GetExtension($status.OutputFile)
                                $contentType = switch ($fileExtension) {
                                    '.md' { 'text/markdown; charset=utf-8' }
                                    '.html' { 'text/html; charset=utf-8' }
                                    '.json' { 'application/json; charset=utf-8' }
                                    '.txt' { 'text/plain; charset=utf-8' }
                                    '.xml' { 'application/xml; charset=utf-8' }
                                    default { 'text/plain; charset=utf-8' }
                                }
                                $response.ContentType = $contentType

                                # For large files (>5MB), use streaming
                                if ($fileInfo.Length -gt 5MB) {
                                    Write-Host "Streaming large file: $($fileInfo.Length) bytes"
                                    $response.SendChunked = $true

                                    try {
                                        $fileStream = [System.IO.File]::OpenRead($status.OutputFile)
                                        $buffer = New-Object byte[] 8192

                                        while (($bytesRead = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                                            $response.OutputStream.Write($buffer, 0, $bytesRead)
                                        }
                                    }
                                    finally {
                                        if ($fileStream) { $fileStream.Close() }
                                    }
                                }
                                else {
                                    # For smaller files, use the original method
                                    $bytes = [System.IO.File]::ReadAllBytes($status.OutputFile)
                                    $response.ContentLength64 = $bytes.Length
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                }

                                $response.OutputStream.Close()
                                $response.Close()
                                continue
                            }
                            elseif ($status.Status -in @('Queued', 'Processing')) {
                                $response.StatusCode = 202
                                $responseContent = @{ status = $status.Status } | ConvertTo-Json
                            }
                            else {
                                $response.StatusCode = 500
                                $responseContent = @{ status = $status.Status; error = $status.Error } | ConvertTo-Json
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                error   = "Failed to serve result"
                                details = $_.Exception.Message
                            } | ConvertTo-Json
                        }
                    }

                    '^/api/documents/(.+)/reset$' {
                        if ($request.HttpMethod -eq 'POST') {
                            $documentId = $Matches[1]
                            try {
                                # Get current document status
                                $allStatus = Get-ProcessingStatus
                                $status = $allStatus[$documentId]

                                if (-not $status) {
                                    $response.StatusCode = 404
                                    $responseContent = @{
                                        error      = "Document not found"
                                        documentId = $documentId
                                    } | ConvertTo-Json
                                }
                                else {
                                    # Reset document to Ready status
                                    Update-ItemStatus $documentId @{
                                        Status       = 'Ready'
                                        Progress     = 0
                                        StartTime    = $null
                                        EndTime      = $null
                                        Error        = $null
                                        ErrorDetails = $null
                                        StdErr       = $null
                                        LastUpdate   = Get-Date
                                    }

                                    $responseContent = @{
                                        success    = $true
                                        documentId = $documentId
                                        status     = 'Ready'
                                        message    = "Document status reset to Ready"
                                    } | ConvertTo-Json
                                }
                            }
                            catch {
                                $response.StatusCode = 500
                                $responseContent = @{
                                    error   = "Failed to reset document status"
                                    details = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                    }

                    default {
                        $response.StatusCode = 404
                        $responseContent = @{ error = "Not found" } | ConvertTo-Json
                    }
                }

                $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                $response.ContentType = "application/json"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.Close()
            }
            catch {
                # Handle any unhandled exceptions in request processing
                Write-Warning "API Server error processing request: $($_.Exception.Message)"
                try {
                    if ($response -and -not $response.OutputStream.CanWrite) {
                        # Response already closed, skip
                        continue
                    }
                    $errorResponse = @{
                        error   = "Internal server error"
                        details = $_.Exception.Message
                    } | ConvertTo-Json
                    $errorBuffer = [System.Text.Encoding]::UTF8.GetBytes($errorResponse)
                    $response.StatusCode = 500
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $errorBuffer.Length
                    $response.OutputStream.Write($errorBuffer, 0, $errorBuffer.Length)
                    $response.Close()
                }
                catch {
                    # If even error handling fails, just continue
                    Write-Warning "Failed to send error response: $($_.Exception.Message)"
                }
            }
        }
    }
    finally {
        $listener.Stop()
    }
}

<#
.SYNOPSIS
    Generates the web frontend files for the PSDocling system.

.DESCRIPTION
    Creates a complete single-page web application including HTML, CSS, and JavaScript files for the PSDocling document processing interface. The frontend provides drag-and-drop file upload, real-time status updates, and download capabilities.

.EXAMPLE
    New-FrontendFiles
    Generates the web frontend files in the DoclingFrontend directory.

.NOTES
    Creates the following files:
    - index.html - Main web interface
    - styles.css - Styling for the interface
    - app.js - JavaScript functionality for file upload and status monitoring

    The frontend is served on port 8081 by default and communicates with the API server on port 8080.
#>
function New-FrontendFiles {
    [CmdletBinding()]
    param()

    $frontendDir = ".\DoclingFrontend"
    if (-not (Test-Path $frontendDir)) {
        New-Item -ItemType Directory -Path $frontendDir -Force | Out-Null
    }

    # Simple HTML file with version
    $version = $script:DoclingSystem.Version
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>PSDocling v$version</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background: linear-gradient(135deg, #1a1a1a 0%, #2d2d2d 100%);
            color: #e0e0e0;
            min-height: 100vh;
        }
        .header {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            padding: 25px;
            border-radius: 12px;
            margin-bottom: 25px;
            border: 1px solid #404040;
            box-shadow: 0 4px 15px rgba(4, 159, 217, 0.1);
        }
        .header h1 {
            margin: 0 0 10px 0;
            background: linear-gradient(45deg, #049fd9, #66d9ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            font-size: 2.2em;
            font-weight: 600;
        }
        .upload-area {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            border: 2px dashed #555;
            border-radius: 12px;
            padding: 40px;
            text-align: center;
            margin-bottom: 25px;
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        .upload-area::before {
            content: '';
            position: absolute;
            top: 0;
            left: -100%;
            width: 100%;
            height: 100%;
            background: linear-gradient(90deg, transparent, rgba(4, 159, 217, 0.1), transparent);
            transition: left 0.5s;
        }
        .upload-area:hover {
            border-color: #049fd9;
            background: linear-gradient(135deg, #2e2e2e 0%, #3e3e3e 100%);
            box-shadow: 0 4px 20px rgba(4, 159, 217, 0.2);
        }
        .upload-area:hover::before {
            left: 100%;
        }
        .btn {
            background: linear-gradient(135deg, #049fd9 0%, #0284c7 100%);
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 500;
            transition: all 0.3s ease;
            box-shadow: 0 2px 10px rgba(4, 159, 217, 0.3);
        }
        .btn:hover {
            background: linear-gradient(135deg, #0284c7 0%, #049fd9 100%);
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(4, 159, 217, 0.4);
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 25px;
        }
        .stat {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            padding: 25px;
            border-radius: 12px;
            text-align: center;
            border: 1px solid #404040;
            transition: all 0.3s ease;
        }
        .stat:hover {
            transform: translateY(-3px);
            box-shadow: 0 6px 20px rgba(4, 159, 217, 0.15);
            border-color: #049fd9;
        }
        .stat-value {
            font-size: 2.5em;
            font-weight: 700;
            background: linear-gradient(45deg, #049fd9, #66d9ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 5px;
        }
        .stat div:last-child {
            color: #b0b0b0;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .results {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            border-radius: 12px;
            padding: 25px;
            border: 1px solid #404040;
        }
        .results h3 {
            color: #049fd9;
            margin-top: 0;
            font-size: 1.3em;
            font-weight: 600;
        }
        .result-item {
            padding: 15px;
            border-bottom: 1px solid #404040;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-radius: 8px;
            margin-bottom: 8px;
            transition: all 0.3s ease;
        }
        .result-item:hover {
            background: rgba(4, 159, 217, 0.1);
            border-color: #049fd9;
        }
        .result-item:last-child {
            border-bottom: none;
            margin-bottom: 0;
        }
        .result-item strong {
            color: #ffffff;
        }
        .result-item a {
            color: #049fd9;
            text-decoration: none;
            padding: 6px 12px;
            border: 1px solid #049fd9;
            border-radius: 6px;
            transition: all 0.3s ease;
            font-size: 0.9em;
        }
        .result-item a:hover {
            background: #049fd9;
            color: white;
            transform: scale(1.05);
        }
        .format-selector {
            background: #1a1a1a;
            border: 1px solid #555;
            color: #e0e0e0;
            padding: 4px 8px;
            border-radius: 4px;
            margin: 0 8px;
            font-size: 0.85em;
        }
        .format-selector:focus {
            border-color: #049fd9;
            outline: none;
        }
        .reprocess-btn {
            background: #666;
            color: white;
            border: none;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.8em;
            cursor: pointer;
            margin-left: 8px;
            transition: background 0.3s ease;
        }
        .reprocess-btn:hover {
            background: #049fd9;
        }
        .hidden { display: none; }
        .progress {
            width: 100%;
            height: 8px;
            background: #404040;
            border-radius: 10px;
            overflow: hidden;
            margin: 15px 0;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #049fd9, #66d9ff);
            width: 0%;
            transition: width 0.3s ease;
            border-radius: 10px;
        }
        #status {
            color: #049fd9;
            font-weight: 600;
        }
        /* Status indicators */
        .status-ready { color: #049fd9; }
        .status-queued { color: #049fd9; }
        .status-processing { color: #049fd9; }
        .status-completed { color: #10b981; }
        .status-error {
            color: #ef4444;
            cursor: pointer;
            text-decoration: underline;
        }
        .status-error:hover {
            color: #f87171;
        }

        /* Progress wheel */
        .progress-container {
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        .progress-wheel {
            width: 16px;
            height: 16px;
            border: 2px solid #404040;
            border-top: 2px solid #049fd9;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            display: inline-block;
            flex-shrink: 0;
        }
        .progress-text {
            font-size: 12px;
            color: #6b7280;
        }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .start-btn {
            background: #3b82f6;
            color: white;
            border: none;
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 0.8em;
            cursor: pointer;
            margin-left: 8px;
            transition: background 0.3s ease;
        }
        .start-btn:hover {
            background: #3b82f6;
        }

        /* Modal styles */
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.7);
        }
        .modal-content {
            background: linear-gradient(135deg, #2a2a2a 0%, #3a3a3a 100%);
            margin: 5% auto;
            padding: 30px;
            border: 1px solid #555;
            border-radius: 12px;
            width: 80%;
            max-width: 800px;
            max-height: 80vh;
            overflow-y: auto;
            color: #e0e0e0;
        }
        .modal-header {
            border-bottom: 1px solid #555;
            padding-bottom: 15px;
            margin-bottom: 20px;
        }
        .modal-header h2 {
            margin: 0;
            color: #ef4444;
        }
        .close {
            color: #aaa;
            float: right;
            font-size: 28px;
            font-weight: bold;
            cursor: pointer;
        }
        .close:hover {
            color: #fff;
        }
        .error-section {
            margin-bottom: 20px;
            padding: 15px;
            background: #1a1a1a;
            border-radius: 8px;
            border-left: 4px solid #ef4444;
        }
        .error-section h3 {
            margin-top: 0;
            color: #f87171;
        }
        .error-code {
            background: #0f0f0f;
            padding: 10px;
            border-radius: 6px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            white-space: pre-wrap;
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid #333;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>PSDocling v$version</h1>
        <p style="margin: 5px 0; color: #b0b0b0; font-size: 1.1em;">PowerShell-based Document Processor</p>
        <p>Backend Status: <span id="status">Connecting...</span></p>
    </div>

    <div class="upload-area" id="drop-zone">
        <h3>Drop files here or click to browse</h3>
        <button class="btn" onclick="document.getElementById('file-input').click()" style="margin: 20px 0;">Choose Files</button>
        <input type="file" id="file-input" multiple accept=".pdf,.docx,.xlsx,.pptx,.md,.html,.xhtml,.csv,.png,.jpg,.jpeg,.tiff,.bmp,.webp" style="display: none;">

        <div style="margin: 25px 0 15px 0; text-align: center;">
            <h4 style="margin: 0 0 8px 0; color: #049fd9; font-size: 1.1em; font-weight: bold;">Supported File Types</h4>
            <p style="margin: 0; color: #b0b0b0; font-size: 0.95em;">PDF, DOCX, XLSX, PPTX, MD, HTML, XHTML, CSV, PNG, JPEG, TIFF, BMP, WEBP</p>
        </div>
    </div>

    <div id="upload-progress" class="hidden">
        <p>Uploading files...</p>
        <div class="progress"><div class="progress-bar" id="progress-bar"></div></div>
    </div>

    <div class="stats">
        <div class="stat"><div class="stat-value" id="queued">0</div><div>Queued</div></div>
        <div class="stat"><div class="stat-value" id="processing">0</div><div>Processing</div></div>
        <div class="stat"><div class="stat-value" id="completed">0</div><div>Completed</div></div>
        <div class="stat"><div class="stat-value" id="errors">0</div><div>Errors</div></div>
    </div>

    <div class="results">
        <h3>Processing Results</h3>
        <div id="results-list"></div>
    </div>

    <div class="results" style="margin-top: 25px;">
        <h3>Processed Files</h3>
        <div id="files-list">
            <p style="color: #b0b0b0; font-style: italic;">Loading processed files...</p>
        </div>
    </div>

    <!-- Error Details Modal -->
    <div id="errorModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <span class="close">&times;</span>
                <h2>Error Details</h2>
            </div>
            <div id="errorModalContent">
                <p>Loading error details...</p>
            </div>
        </div>
    </div>

    <script>
    const API = 'http://localhost:8080';
    const results = {};

    document.addEventListener('DOMContentLoaded', function() {
        setupUpload();
        // Delay initial API calls to give server time to start
        setTimeout(async () => {
            const isHealthy = await checkHealth();
            if (isHealthy) {
                loadExistingDocuments();
                loadProcessedFiles();
            }
        }, 1000);
        setInterval(loadProcessedFiles, 10000); // Refresh processed files every 10 seconds
        setInterval(updateStats, 2000);
    });

    function setupUpload() {
        const zone = document.getElementById('drop-zone');
        const input = document.getElementById('file-input');

        zone.addEventListener('dragover', e => { e.preventDefault(); zone.style.borderColor = '#007cba'; });
        zone.addEventListener('dragleave', () => { zone.style.borderColor = '#ccc'; });
        zone.addEventListener('drop', e => { e.preventDefault(); zone.style.borderColor = '#ccc'; handleFiles(e.dataTransfer.files); });
        input.addEventListener('change', e => handleFiles(e.target.files));
    }

    async function handleFiles(files) {
        const progress = document.getElementById('upload-progress');
        const bar = document.getElementById('progress-bar');

        progress.classList.remove('hidden');

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                // Use FileReader for reliable base64 conversion
                const base64 = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => {
                        const result = reader.result;
                        const base64Data = result.substring(result.indexOf(',') + 1);
                        resolve(base64Data);
                    };
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                });

                // Default enrichment options for initial upload (user selects them later)
                const enrichCode = false;
                const enrichFormula = false;
                const enrichPictureClasses = false;
                const enrichPictureDescription = false;

                const response = await fetch(API + '/api/upload', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        fileName: file.name,
                        dataBase64: base64,
                        enrichCode: enrichCode,
                        enrichFormula: enrichFormula,
                        enrichPictureClasses: enrichPictureClasses,
                        enrichPictureDescription: enrichPictureDescription
                    })
                });

                if (response.ok) {
                    const result = await response.json();
                    addResult(result.documentId, file.name);
                    // Don't start polling - document should stay in Ready status for manual conversion
                } else {
                    throw new Error('Upload failed');
                }
            } catch (error) {
                alert('Error uploading ' + file.name + ': ' + error.message);
            }

            bar.style.width = ((i + 1) / files.length * 100) + '%';
        }

        setTimeout(() => { progress.classList.add('hidden'); bar.style.width = '0%'; }, 1000);
    }

    // Validate export format selection
    function validateExportFormat(id) {
        const radioButtons = document.querySelectorAll('input[name="format-' + id + '"]');
        const validationMsg = document.getElementById('validation-' + id);
        const startBtn = document.getElementById('start-' + id);

        let selectedCount = 0;
        radioButtons.forEach(radio => {
            if (radio.checked) selectedCount++;
        });

        if (selectedCount === 1) {
            validationMsg.style.display = 'none';
            startBtn.disabled = false;
            startBtn.style.opacity = '1';
            startBtn.style.cursor = 'pointer';
        } else {
            validationMsg.style.display = 'block';
            startBtn.disabled = true;
            startBtn.style.opacity = '0.5';
            startBtn.style.cursor = 'not-allowed';
        }
    }

    function addResult(id, name, currentFormat = 'markdown') {
        const list = document.getElementById('results-list');
        const item = document.createElement('div');
        item.className = 'result-item';
        item.innerHTML =
            '<div style="width: 100%;">' +
                '<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">' +
                    '<strong>' + name + '</strong>' +
                    '<span id="status-' + id + '" class="status-ready">Ready</span>' +
                '</div>' +
                '<div id="validation-' + id + '" style="color: #ef4444; font-size: 0.85em; margin-bottom: 10px; display: none;">&#9888; Select a single export format</div>' +
                '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 15px;">' +
                    '<div style="padding: 15px; background: #2a2a2a; border-radius: 8px; border: 1px solid #404040;">' +
                        '<h4 style="margin: 0 0 10px 0; color: #049fd9; font-size: 1em; font-weight: bold;">Export Formats</h4>' +
                        '<div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.9em;">' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="radio" name="format-' + id + '" value="markdown"' + (currentFormat === 'markdown' ? ' checked' : '') + ' style="margin: 0;" onchange="validateExportFormat(\'' + id + '\')">' +
                                '<span>Markdown (.md)</span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="radio" name="format-' + id + '" value="html"' + (currentFormat === 'html' ? ' checked' : '') + ' style="margin: 0;" onchange="validateExportFormat(\'' + id + '\')">' +
                                '<span>HTML (.html)</span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="radio" name="format-' + id + '" value="json"' + (currentFormat === 'json' ? ' checked' : '') + ' style="margin: 0;" onchange="validateExportFormat(\'' + id + '\')">' +
                                '<span>JSON (.json)</span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="radio" name="format-' + id + '" value="text"' + (currentFormat === 'text' ? ' checked' : '') + ' style="margin: 0;" onchange="validateExportFormat(\'' + id + '\')">' +
                                '<span>Plain Text (.txt)</span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="radio" name="format-' + id + '" value="doctags"' + (currentFormat === 'doctags' ? ' checked' : '') + ' style="margin: 0;" onchange="validateExportFormat(\'' + id + '\')">' +
                                '<span>DocTags (.xml)</span>' +
                            '</label>' +
                        '</div>' +
                        '<div style="margin-top: 12px; padding-top: 10px; border-top: 1px solid #404040;">' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 0.9em;">' +
                                '<input type="checkbox" id="embedImages-' + id + '" style="margin: 0;">' +
                                '<span>Embed Images</span>' +
                            '</label>' +
                        '</div>' +
                    '</div>' +
                    '<div style="padding: 15px; background: #2a2a2a; border-radius: 8px; border: 1px solid #404040;">' +
                        '<h4 style="margin: 0 0 10px 0; color: #049fd9; font-size: 1em; font-weight: bold;">Enrichment Options</h4>' +
                        '<div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.9em;">' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="checkbox" id="enrichCode-' + id + '" style="margin: 0;">' +
                                '<span>Code Understanding <span style="color: #049fd9; font-size: 0.8em;">(Requires Windows Developer Mode)</span></span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="checkbox" id="enrichFormula-' + id + '" style="margin: 0;">' +
                                '<span>Formula Understanding <span style="color: #049fd9; font-size: 0.8em;">(Downloads CodeFormulaV2 model)</span></span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="checkbox" id="enrichPictureClasses-' + id + '" style="margin: 0;">' +
                                '<span>Picture Classification <span style="color: #049fd9; font-size: 0.8em;">(Downloads classification models)</span></span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="checkbox" id="enrichPictureDescription-' + id + '" style="margin: 0;">' +
                                '<span>Picture Description <span style="color: #049fd9; font-size: 0.8em;">(Downloads Granite Vision model)</span></span>' +
                            '</label>' +
                        '</div>' +
                    '</div>' +
                '</div>' +
                '<div style="display: flex; justify-content: space-between; align-items: center;">' +
                    '<div>' +
                        '<button class="start-btn" onclick="startConversion(\'' + id + '\')" id="start-' + id + '" disabled>Start Conversion</button>' +
                        '<button class="reprocess-btn" onclick="reprocessDocument(\'' + id + '\')" style="display:none" id="reprocess-' + id + '">Re-process</button>' +
                    '</div>' +
                    '<a id="link-' + id + '" href="#" style="display:none">Download</a>' +
                '</div>' +
            '</div>';
        list.appendChild(item);
        results[id] = { name: name, format: currentFormat };

        // Initialize validation for the new item
        validateExportFormat(id);
    }

    // Start conversion for a ready document
    async function startConversion(id) {
        // Get selected format from radio buttons
        const radioButtons = document.querySelectorAll('input[name="format-' + id + '"]');
        let selectedFormat = null;
        radioButtons.forEach(radio => {
            if (radio.checked) selectedFormat = radio.value;
        });

        if (!selectedFormat) {
            alert('Please select an export format');
            return;
        }
        const embedImagesCheckbox = document.getElementById('embedImages-' + id);
        const embedImages = embedImagesCheckbox.checked;

        // Get enrichment options
        const enrichCode = document.getElementById('enrichCode-' + id).checked;
        const enrichFormula = document.getElementById('enrichFormula-' + id).checked;
        const enrichPictureClasses = document.getElementById('enrichPictureClasses-' + id).checked;
        const enrichPictureDescription = document.getElementById('enrichPictureDescription-' + id).checked;

        const statusElement = document.getElementById('status-' + id);
        const startBtn = document.getElementById('start-' + id);

        try {
            // Hide start button and update status
            startBtn.style.display = 'none';
            statusElement.textContent = 'Queued...';
            statusElement.className = 'status-queued';

            const response = await fetch(API + '/api/start-conversion', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    documentId: id,
                    exportFormat: selectedFormat,
                    embedImages: embedImages,
                    enrichCode: enrichCode,
                    enrichFormula: enrichFormula,
                    enrichPictureClasses: enrichPictureClasses,
                    enrichPictureDescription: enrichPictureDescription
                })
            });

            if (response.ok) {
                const result = await response.json();

                // Update the stored format
                if (results[id] && typeof results[id] === 'object') {
                    results[id].format = selectedFormat;
                } else {
                    results[id] = { name: results[id] || 'Unknown', format: selectedFormat };
                }

                // Start polling for completion (with immediate first check for fast documents)
                setTimeout(() => pollResult(id, results[id].name || 'Unknown'), 500);
            } else {
                throw new Error('Failed to start conversion');
            }
        } catch (error) {
            statusElement.textContent = 'Error (click for details)';
            statusElement.className = 'status-error';
            statusElement.onclick = () => showErrorDetails(id, results[id].name || 'Unknown');
            startBtn.style.display = 'inline'; // Show start button again
            console.error('Start conversion error:', error);
        }
    }

    // Re-process document with new format
    async function reprocessDocument(id) {
        // Get selected format from radio buttons
        const radioButtons = document.querySelectorAll('input[name="format-' + id + '"]');
        let newFormat = null;
        radioButtons.forEach(radio => {
            if (radio.checked) newFormat = radio.value;
        });

        if (!newFormat) {
            alert('Please select an export format');
            return;
        }
        const embedImagesCheckbox = document.getElementById('embedImages-' + id);
        const embedImages = embedImagesCheckbox.checked;

        // Get enrichment options
        const enrichCode = document.getElementById('enrichCode-' + id).checked;
        const enrichFormula = document.getElementById('enrichFormula-' + id).checked;
        const enrichPictureClasses = document.getElementById('enrichPictureClasses-' + id).checked;
        const enrichPictureDescription = document.getElementById('enrichPictureDescription-' + id).checked;

        const statusElement = document.getElementById('status-' + id);
        const reprocessBtn = document.getElementById('reprocess-' + id);

        try {
            // Hide reprocess button and update status
            reprocessBtn.style.display = 'none';
            statusElement.textContent = 'Re-processing...';
            statusElement.className = 'status-processing';

            const response = await fetch(API + '/api/reprocess', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    documentId: id,
                    exportFormat: newFormat,
                    embedImages: embedImages,
                    enrichCode: enrichCode,
                    enrichFormula: enrichFormula,
                    enrichPictureClasses: enrichPictureClasses,
                    enrichPictureDescription: enrichPictureDescription
                })
            });

            if (response.ok) {
                const result = await response.json();

                // Update the stored format
                if (results[id] && typeof results[id] === 'object') {
                    results[id].format = newFormat;
                } else {
                    results[id] = { name: results[id] || 'Unknown', format: newFormat };
                }

                // Start polling for completion (with immediate first check for fast documents)
                setTimeout(() => pollResult(id, results[id].name || 'Unknown'), 500);
            } else {
                throw new Error('Failed to reprocess document');
            }
        } catch (error) {
            statusElement.textContent = 'Error (click for details)';
            statusElement.className = 'status-error';
            statusElement.onclick = () => showErrorDetails(id, results[id].name || 'Unknown');
            console.error('Reprocess error:', error);
        }
    }

    // Error modal functions
    function showErrorDetails(id, fileName) {
        const modal = document.getElementById('errorModal');
        const content = document.getElementById('errorModalContent');

        content.innerHTML = '<p>Loading error details...</p>';
        modal.style.display = 'block';

        fetch(API + '/api/error/' + id)
            .then(response => {
                console.log('Error API response status:', response.status);
                if (response.ok) {
                    return response.json();
                } else {
                    return response.text().then(text => {
                        console.log('Error API response text:', text);
                        throw new Error('HTTP ' + response.status + ': ' + text);
                    });
                }
            })
            .then(errorData => {
                console.log('Error data received:', errorData);
                content.innerHTML = formatErrorDetails(errorData);
            })
            .catch(error => {
                console.error('Error loading error details:', error);
                content.innerHTML = '<p style="color: #ef4444;">Failed to load error details: ' + error.message + '</p>';
            });
    }

    function formatEstimatedTime(estimatedDurationMs, elapsedTimeMs) {
        if (!estimatedDurationMs || !elapsedTimeMs) return '';

        const remainingMs = Math.max(0, estimatedDurationMs - elapsedTimeMs);
        const remainingSeconds = Math.round(remainingMs / 1000);

        if (remainingSeconds <= 0) return 'finishing...';
        if (remainingSeconds < 60) return remainingSeconds + 's remaining';

        const remainingMinutes = Math.floor(remainingSeconds / 60);
        const seconds = remainingSeconds % 60;

        if (remainingMinutes < 60) {
            return seconds > 0 ? remainingMinutes + 'm ' + seconds + 's remaining' : remainingMinutes + 'm remaining';
        }

        const hours = Math.floor(remainingMinutes / 60);
        const minutes = remainingMinutes % 60;
        return minutes > 0 ? hours + 'h ' + minutes + 'm remaining' : hours + 'h remaining';
    }


    function formatErrorDetails(errorData) {
        let html = '<div class="error-section">';
        html += '<h3>File Information</h3>';
        html += '<p><strong>File:</strong> ' + (errorData.fileName || 'Unknown') + '</p>';
        html += '<p><strong>Document ID:</strong> ' + errorData.id + '</p>';
        html += '<p><strong>Queued:</strong> ' + (errorData.queuedTime?.DateTime || errorData.queuedTime || 'Unknown') + '</p>';

        // Check if this is status information or error details
        if (errorData.currentStatus && errorData.currentStatus !== 'Error') {
            // This is status information, not error details
            html += '<p><strong>Current Status:</strong> ' + errorData.currentStatus + '</p>';
            if (errorData.startTime) {
                html += '<p><strong>Started:</strong> ' + (errorData.startTime?.DateTime || errorData.startTime || 'Unknown') + '</p>';
            }
            if (errorData.progress !== undefined) {
                html += '<p><strong>Progress:</strong> ' + errorData.progress + '%</p>';
            }
            html += '</div>';

            html += '<div class="error-section">';
            html += '<h3>Status Information</h3>';
            html += '<div class="error-code" style="background: #1a2e1a; border-left: 4px solid #10b981;">' + (errorData.message || 'Document is being processed') + '</div>';
            html += '</div>';
        } else {
            // This is actual error details
            html += '<p><strong>Failed:</strong> ' + (errorData.endTime?.DateTime || errorData.endTime || 'Unknown') + '</p>';
            html += '</div>';

            html += '<div class="error-section">';
            html += '<h3>Error Message</h3>';
            html += '<div class="error-code">' + (errorData.error || 'No error message available') + '</div>';
            html += '</div>';
        }

        if (errorData.stderr && typeof errorData.stderr === 'string' && errorData.stderr.trim()) {
            html += '<div class="error-section">';
            html += '<h3>Python Error Output (stderr)</h3>';
            html += '<div class="error-code">' + errorData.stderr + '</div>';
            html += '</div>';
        }

        if (errorData.errorDetails) {
            html += '<div class="error-section">';
            html += '<h3>Technical Details</h3>';

            if (errorData.errorDetails.ExceptionType) {
                html += '<p><strong>Exception Type:</strong> ' + errorData.errorDetails.ExceptionType + '</p>';
            }

            if (errorData.errorDetails.InnerException) {
                html += '<p><strong>Inner Exception:</strong> ' + errorData.errorDetails.InnerException + '</p>';
            }

            if (errorData.errorDetails.StackTrace) {
                html += '<h4>Stack Trace:</h4>';
                html += '<div class="error-code">' + errorData.errorDetails.StackTrace + '</div>';
            }

            if (errorData.errorDetails.ScriptStackTrace) {
                html += '<h4>Script Stack Trace:</h4>';
                html += '<div class="error-code">' + errorData.errorDetails.ScriptStackTrace + '</div>';
            }

            html += '</div>';
        }

        return html;
    }

    // Setup modal close functionality
    document.addEventListener('DOMContentLoaded', function() {
        const modal = document.getElementById('errorModal');
        const closeBtn = document.querySelector('.close');

        closeBtn.onclick = function() {
            modal.style.display = 'none';
        }

        window.onclick = function(event) {
            if (event.target === modal) {
                modal.style.display = 'none';
            }
        }
    });

    async function pollResult(id, name, attempt = 0) {
        try {
            const response = await fetch(API + '/api/result/' + id);
            if (response.status === 200) {
                const contentLength = response.headers.get('content-length');
                const blob = await response.blob();

                document.getElementById('status-' + id).textContent = 'Completed';
                const link = document.getElementById('link-' + id);

                // For large files (>5MB) or JSON files >1MB, use download instead of blob URL
                if ((contentLength && parseInt(contentLength) > 5 * 1024 * 1024) ||
                    (blob.type.includes('json') && blob.size > 1 * 1024 * 1024)) {
                    // Use download link instead of blob URL to avoid browser memory issues
                    link.href = API + '/api/result/' + id;
                    link.download = name + '.' + (blob.type.includes('json') ? 'json' : 'md');
                    link.textContent = 'Download (' + (blob.size / (1024 * 1024)).toFixed(1) + ' MB)';
                } else {
                    const url = URL.createObjectURL(blob);
                    link.href = url;
                }

                link.style.display = 'inline';

                // Refresh the Processed Files section immediately when document completes
                loadProcessedFiles();

                // Force an immediate page refresh to ensure all updates are reflected
                window.location.reload();

                return;
            }
            if (response.status === 202) {
                // Check if we can get updated status with progress
                try {
                    const documentsResponse = await fetch(API + '/api/documents');
                    if (documentsResponse.ok) {
                        const documents = await documentsResponse.json();
                        const doc = documents.find(d => d.id === id);
                        if (doc && doc.status === 'Processing') {
                            const statusElement = document.getElementById('status-' + id);
                            let progressText = 'Processing...';
                            if (doc.progress !== undefined && doc.progress !== null) {
                                progressText = 'Processing ' + doc.progress + '%';
                            }

                            statusElement.innerHTML = '<div class="progress-container">' +
                                '<div class="progress-wheel"></div>' +
                                '<span>' + progressText + '</span>' +
                            '</div>';
                        }
                    }
                } catch (e) {
                    // Fallback to simple text if API call fails
                    document.getElementById('status-' + id).textContent = 'Processing...';
                }
                setTimeout(() => pollResult(id, name, attempt + 1), 1000);
                return;
            }

            // Before marking as error, check if document is actually in error state
            try {
                const documentsResponse = await fetch(API + '/api/documents');
                if (documentsResponse.ok) {
                    const documents = await documentsResponse.json();
                    const doc = documents.find(d => d.id === id);
                    if (doc && doc.status === 'Error') {
                        // Document is actually in error state
                        const statusElement = document.getElementById('status-' + id);
                        statusElement.textContent = 'Error (click for details)';
                        statusElement.className = 'status-error';
                        statusElement.onclick = () => showErrorDetails(id, name);
                        return;
                    } else if (doc) {
                        // Document is not in error state, continue polling
                        setTimeout(() => pollResult(id, name, attempt + 1), 2000);
                        return;
                    }
                }
            } catch (docError) {
                console.log('Failed to check document status:', docError);
            }

            // Fallback: mark as connection error but continue trying
            const statusElement = document.getElementById('status-' + id);
            statusElement.textContent = 'Connection Error - Retrying...';
            statusElement.className = 'status-error';
            statusElement.onclick = null; // Don't show error details for connection errors
        } catch (error) {
            if (attempt < 30) {
                setTimeout(() => pollResult(id, name, attempt + 1), 2000);
            } else {
                // After many retries, check if document is actually in error state
                try {
                    const documentsResponse = await fetch(API + '/api/documents');
                    if (documentsResponse.ok) {
                        const documents = await documentsResponse.json();
                        const doc = documents.find(d => d.id === id);
                        if (doc && doc.status === 'Error') {
                            // Document is actually in error state
                            const statusElement = document.getElementById('status-' + id);
                            statusElement.textContent = 'Error (click for details)';
                            statusElement.className = 'status-error';
                            statusElement.onclick = () => showErrorDetails(id, name);
                            return;
                        }
                    }
                } catch (docError) {
                    console.log('Failed to check document status after retries:', docError);
                }

                // Final fallback: assume connection issues
                const statusElement = document.getElementById('status-' + id);
                statusElement.textContent = 'Connection Lost';
                statusElement.className = 'status-error';
                statusElement.onclick = null; // Don't show error details for connection errors
            }
        }
    }

    async function updateStats() {
        try {
            const response = await fetch(API + '/api/status');
            if (response.ok) {
                const stats = await response.json();
                document.getElementById('queued').textContent = stats.QueuedCount || 0;
                document.getElementById('processing').textContent = stats.ProcessingCount || 0;
                document.getElementById('completed').textContent = stats.CompletedCount || 0;
                document.getElementById('errors').textContent = stats.ErrorCount || 0;
            } else {
                console.error('Stats update failed with status:', response.status);
                // Try to reconnect if stats fail
                setTimeout(() => checkHealth(1), 1000);
            }
        } catch (error) {
            console.error('Stats update failed:', error);
            // Try to reconnect if stats fail
            setTimeout(() => checkHealth(1), 1000);
        }
    }

    async function loadExistingDocuments() {
        try {
            const response = await fetch(API + '/api/documents');
            if (response.ok) {
                const documents = await response.json();
                const list = document.getElementById('results-list');
                list.innerHTML = ''; // Clear existing items

                documents.forEach(doc => {
                    console.log('Processing doc in updateDisplay:', doc.id, doc.status);

                    // Skip completed documents - they will be handled by loadCompletedDocuments
                    if (doc.status === 'Completed') {
                        return;
                    }

                    const docFormat = doc.exportFormat || 'markdown';
                    addResult(doc.id, doc.fileName, docFormat);
                    results[doc.id] = { name: doc.fileName, format: docFormat };

                    // Set appropriate status and start polling if needed
                    const statusElement = document.getElementById('status-' + doc.id);

                    if (doc.status === 'Ready') {
                        statusElement.textContent = 'Ready';
                        statusElement.className = 'status-ready';
                        statusElement.onclick = null; // Clear any existing error click handler
                        // Show start button for ready items
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) {
                            startBtn.style.display = 'inline';
                        }
                    } else if (doc.status === 'Processing') {
                        console.log('Processing document:', doc.id, 'with progress:', doc.progress);

                        // Display progress percentage if available
                        let progressText = 'Processing...';
                        if (doc.progress !== undefined && doc.progress !== null) {
                            progressText = 'Processing ' + doc.progress + '%';
                        }

                        statusElement.innerHTML = '<div class="progress-container">' +
                            '<div class="progress-wheel"></div>' +
                            '<span>' + progressText + '</span>' +
                        '</div>';
                        statusElement.className = 'status-processing';
                        statusElement.onclick = null; // Clear any existing error click handler

                        // Hide start button during processing
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) {
                            startBtn.style.display = 'none';
                        }
                        pollResult(doc.id, doc.fileName);
                    } else if (doc.status === 'Queued') {
                        statusElement.textContent = 'Queued...';
                        statusElement.className = 'status-queued';
                        statusElement.onclick = null; // Clear any existing error click handler
                        // Hide start button when queued
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) {
                            startBtn.style.display = 'none';
                        }
                        pollResult(doc.id, doc.fileName);
                    } else if (doc.status === 'Error') {
                        statusElement.textContent = 'Error (click for details)';
                        statusElement.className = 'status-error';
                        statusElement.onclick = () => showErrorDetails(doc.id, doc.fileName);
                        // Show start button again for retry
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) {
                            startBtn.style.display = 'inline';
                        }
                    } else {
                        // CATCH-ALL: Show any unknown status
                        console.log('UNKNOWN STATUS:', doc.status);
                        statusElement.innerHTML = '<span style=\"color: orange; font-weight: bold;\">STATUS: ' + (doc.status || 'UNDEFINED') + '</span>';
                    }
                });
            }
        } catch (error) {
            console.error('Failed to load existing documents:', error);
        }
    }

    async function loadProcessedFiles() {
        try {
            // Load both static files and completed documents
            const [filesResponse, documentsResponse] = await Promise.all([
                fetch(API + '/api/files'),
                fetch(API + '/api/documents')
            ]);

            // If both calls failed, surface a connection message and bail early
            if (!filesResponse.ok && !documentsResponse.ok) {
                console.error('Both /api/files and /api/documents failed.');
                const filesList = document.getElementById('files-list');
                filesList.innerHTML = '<p style="color: #fbbf24;">Server responded with an error. Checking connection...</p>';
                setTimeout(() => checkHealth(1), 2000);
                return;
            }

            const allItems = [];

            // Get documents for document ID mapping
            let documentsMap = new Map();
            if (documentsResponse.ok) {
                const documents = await documentsResponse.json();
                documents
                    .filter(doc => doc.status === 'Completed')
                    .forEach(doc => {
                        documentsMap.set(doc.id, doc);
                    });
            }

            // Add only generated files (output files, not original uploads)
            if (filesResponse.ok) {
                let files = await filesResponse.json();
                if (!Array.isArray(files)) files = [files];

                files.forEach(file => {
                    // Only show generated files, not original uploaded files
                    // Generated files have extensions like .md, .xml, .html, .json
                    const isGeneratedFile = /\.(md|xml|html|json)$/i.test(file.fileName);

                    if (isGeneratedFile) {
                        // Find corresponding document for re-process functionality
                        const correspondingDoc = documentsMap.get(file.id);

                        allItems.push({
                            type: 'file',
                            id: file.id,
                            fileName: file.fileName,
                            fileSize: file.fileSize,
                            lastModified: file.lastModified,
                            downloadUrl: file.downloadUrl,
                            exportFormat: correspondingDoc ? correspondingDoc.exportFormat : 'unknown',
                            canReprocess: !!correspondingDoc
                        });
                    }
                });
            }

            const filesList = document.getElementById('files-list');

            if (allItems.length === 0) {
                filesList.innerHTML = '<p style="color: #b0b0b0; font-style: italic;">No processed files found.</p>';
                return;
            }

            filesList.innerHTML = allItems.map(item => {
                return '<div class="result-item">' +
                    '<div>' +
                    '<strong>' + item.fileName + '</strong><br>' +
                    '<small style="color: #b0b0b0;">' +
                    'Size: ' + item.fileSize + ' | Modified: ' + item.lastModified +
                    '</small>' +
                    '</div>' +
                    '<div>' +
                    '<a href="' + API + item.downloadUrl + '" target="_blank">Download</a>' +
                    (item.canReprocess ?
                        '<button class="reprocess-btn" onclick="reprocessFromCompleted(\'' + item.id + '\', \'' + item.fileName + '\')" style="margin-left: 8px;">Re-process</button>' :
                        '') +
                    '</div>' +
                    '</div>';
            }).join('');
        } catch (error) {
            console.error('Failed to load processed files:', error);
            document.getElementById('files-list').innerHTML =
                '<p style="color: #fbbf24;">Connection lost. Attempting to reconnect...</p>';
            setTimeout(() => checkHealth(1), 2000);
        }
    }

    async function checkHealth(retries = 3) {
        try {
            // Create a timeout promise that rejects after 5 seconds
            const timeoutPromise = new Promise((_, reject) => {
                setTimeout(() => reject(new Error('Timeout')), 5000);
            });

            // Race the fetch against the timeout
            const response = await Promise.race([
                fetch(API + '/api/health'),
                timeoutPromise
            ]);

            if (response.ok) {
                const data = await response.json();
                document.getElementById('status').textContent = 'Connected';
                document.getElementById('status').style.color = '#10b981'; // Green color
                // If we just reconnected, refresh the processed files
                if (document.getElementById('files-list').innerHTML.includes('Connection lost') ||
                    document.getElementById('files-list').innerHTML.includes('Server responded with an error')) {
                    loadProcessedFiles();
                }
                return true;
            } else {
                document.getElementById('status').textContent = 'Server Error';
                document.getElementById('status').style.color = '#ef4444'; // Red color
                return false;
            }
        } catch (error) {
            console.log('Health check error:', error.message);
            if (retries > 0) {
                document.getElementById('status').textContent = 'Connecting...';
                document.getElementById('status').style.color = '#fbbf24'; // Yellow color
                await new Promise(resolve => setTimeout(resolve, 2000));
                return checkHealth(retries - 1);
            } else {
                document.getElementById('status').textContent = 'Disconnected';
                document.getElementById('status').style.color = '#ef4444'; // Red color
                return false;
            }
        }
    }

    // Function to move a completed document back to Processing Results for re-processing
    async function reprocessFromCompleted(documentId, fileName) {
        try {
            // Get the document details to restore it to Processing Results
            const documentsResponse = await fetch(API + '/api/documents');
            if (documentsResponse.ok) {
                const documents = await documentsResponse.json();
                const doc = documents.find(d => d.id === documentId);

                if (doc) {
                    // Add the document back to Processing Results with current settings
                    const docFormat = doc.exportFormat || 'markdown';
                    addResult(documentId, fileName, docFormat);
                    results[documentId] = { name: fileName, format: docFormat };

                    // Set status to Ready so user can configure options
                    const statusElement = document.getElementById('status-' + documentId);
                    statusElement.textContent = 'Ready';
                    statusElement.className = 'status-ready';
                    statusElement.onclick = null;

                    // Show start button
                    const startBtn = document.getElementById('start-' + documentId);
                    if (startBtn) {
                        startBtn.style.display = 'inline';
                    }

                    // Update the document status to Ready in the backend
                    await fetch(API + '/api/documents/' + documentId + '/reset', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ status: 'Ready' })
                    });

                    // Refresh both sections to ensure proper display
                    loadProcessedFiles();

                    // Scroll to the Processing Results section so user can see the document
                    const processingSection = document.getElementById('results');
                    if (processingSection) {
                        processingSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                } else {
                    throw new Error('Document not found');
                }
            } else {
                throw new Error('Failed to get document details');
            }
        } catch (error) {
            alert('Error moving document back to Processing Results: ' + error.message);
        }
    }
    </script>
</body>
</html>
"@

    $html | Set-Content (Join-Path $frontendDir "index.html") -Encoding UTF8

    # Web server script
    $webServer = @'
param([int]$Port = 8081)

$http = New-Object System.Net.HttpListener
$http.Prefixes.Add("http://localhost:$Port/")
$http.Start()

Write-Host "Web server running at http://localhost:$Port" -ForegroundColor Green

try {
    while ($http.IsListening) {
        $context = $http.GetContext()
        $response = $context.Response

        $path = $context.Request.Url.LocalPath
        if ($path -eq "/") { $path = "/index.html" }

        $filePath = Join-Path $PSScriptRoot $path.TrimStart('/')

        if (Test-Path $filePath) {
            $content = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentType = "text/html"
            $response.ContentLength64 = $content.Length
            $response.OutputStream.Write($content, 0, $content.Length)
        } else {
            $response.StatusCode = 404
        }

        $response.Close()
    }
} finally {
    $http.Stop()
}
'@

    $webServer | Set-Content (Join-Path $frontendDir "Start-WebServer.ps1") -Encoding UTF8

    Write-Host "Frontend files created in $frontendDir" -ForegroundColor Green
}

<#
.SYNOPSIS
    Starts all components of the PSDocling system.

.DESCRIPTION
    Launches all three core components of the PSDocling system: the API server, document processor, and web frontend. Each component runs in a separate background job for concurrent operation.

.PARAMETER OpenBrowser
    When specified, automatically opens the web frontend in the default browser after starting all services.

.EXAMPLE
    Start-DoclingSystem
    Starts all PSDocling services without opening the browser.

.EXAMPLE
    Start-DoclingSystem -OpenBrowser
    Starts all PSDocling services and opens the web interface in the default browser.

.NOTES
    This function starts:
    - API Server (default port 8080) in background job
    - Document Processor in background job
    - Web Frontend server (default port 8081) in background job

    Use Get-DoclingSystemStatus to monitor the system status.
#>
function Start-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$OpenBrowser
    )

    Write-Host "Starting Docling System..." -ForegroundColor Cyan

    # Start API server
    $apiScript = "Import-Module '$PSCommandPath' -Force; Start-APIServer -Port $($script:DoclingSystem.APIPort)"
    $apiPath = Join-Path $env:TEMP "docling_api.ps1"
    $apiScript | Set-Content $apiPath -Encoding UTF8

    $apiProcess = Start-Process powershell -ArgumentList "-File", $apiPath -PassThru -WindowStyle Hidden
    Write-Host "API server started on port $($script:DoclingSystem.APIPort)" -ForegroundColor Green

    # Start processor
    $procScript = "Import-Module '$PSCommandPath' -Force; Start-DocumentProcessor"
    $procPath = Join-Path $env:TEMP "docling_processor.ps1"
    $procScript | Set-Content $procPath -Encoding UTF8

    $procProcess = Start-Process powershell -ArgumentList "-File", $procPath -PassThru -WindowStyle Hidden
    Write-Host "Document processor started" -ForegroundColor Green

    # Start web server
    $webPath = ".\DoclingFrontend\Start-WebServer.ps1"
    if (Test-Path $webPath) {
        $webProcess = Start-Process powershell -ArgumentList "-File", $webPath, "-Port", $script:DoclingSystem.WebPort -PassThru -WindowStyle Hidden
        Write-Host "Web server started on port $($script:DoclingSystem.WebPort)" -ForegroundColor Green

        if ($OpenBrowser) {
            Start-Sleep 2
            Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
        }
    }

    Write-Host "System running! Frontend: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green

    return @{
        API       = $apiProcess
        Processor = $procProcess
        Web       = $webProcess
    }
}

<#
.SYNOPSIS
    Gets the current status of the PSDocling system.

.DESCRIPTION
    Retrieves comprehensive status information about the PSDocling system including queue status, processing statistics, Python availability, and background job information.

.EXAMPLE
    Get-DoclingSystemStatus
    Returns a status object with queue count, processing statistics, and system information.

.OUTPUTS
    PSCustomObject with the following properties:
    - QueueCount: Number of documents waiting to be processed
    - ProcessingCount: Number of documents currently being processed
    - CompletedCount: Number of documents completed
    - PythonAvailable: Boolean indicating if Python and Docling are available
    - BackgroundJobs: Information about running background jobs

.NOTES
    This function provides real-time status information useful for monitoring the system health and processing progress.
#>
function Get-DoclingSystemStatus {
    $queue = Get-QueueItems
    $allStatus = Get-ProcessingStatus
    $processing = $allStatus.Values | Where-Object { $_.Status -eq 'Processing' }
    $completed = $allStatus.Values | Where-Object { $_.Status -eq 'Completed' }

    # Test API
    $apiHealthy = $false
    try {
        $response = Invoke-WebRequest -Uri "http://localhost:$($script:DoclingSystem.APIPort)/api/health" -UseBasicParsing -TimeoutSec 2 -ErrorAction SilentlyContinue
        $apiHealthy = $response.StatusCode -eq 200
    }
    catch {}

    return @{
        Initialized = $true
        Backend     = @{
            Running          = $true
            ProcessorRunning = $true
            APIHealthy       = $apiHealthy
            QueueCount       = $queue.Count
            ProcessingCount  = @($processing).Count
        }
        Frontend    = @{
            Running = $true
            Port    = $script:DoclingSystem.WebPort
            URL     = "http://localhost:$($script:DoclingSystem.WebPort)"
        }
        System      = @{
            Version                 = $script:DoclingSystem.Version
            TotalDocumentsProcessed = @($completed).Count
        }
    }
}

<#
.SYNOPSIS
    Gets the Python and Docling library availability status.

.DESCRIPTION
    Returns a boolean value indicating whether Python and the Docling library are available and properly configured for document processing.

.EXAMPLE
    Get-PythonStatus
    Returns $true if Python and Docling are available, $false otherwise.

.OUTPUTS
    Boolean value indicating Python/Docling availability.

.NOTES
    This status is determined during module load and system initialization. If Python is not available, document processing will not function.
#>
function Get-PythonStatus {
    return $script:DoclingSystem.PythonAvailable
}

Export-ModuleMember -Function @(
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

Write-Host "Docling System v$($script:DoclingSystem.Version) loaded successfully!" -ForegroundColor Green