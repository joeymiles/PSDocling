<#
.SYNOPSIS
    Start-DocumentProcessor function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
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

    # Use same directory as output file for images
    images_dir = dst.parent
    # No need to create directory as it already exists for the output file

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

                            # Create reference in content (just filename since in same folder)
                            relative_path = image_filename
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
                        # Images are now in same folder as output file
                        $statusUpdate.ImagesDirectory = Split-Path $outputFile -Parent
                        Write-Host "Completed: $($item.FileName) ($imagesExtracted images extracted)" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Completed: $($item.FileName)" -ForegroundColor Green
                    }

                    # Process chunking if enabled
                    if ($item.EnableChunking -and $outputFile) {
                        try {
                            Write-Host "Starting hybrid chunking for $($item.FileName)..." -ForegroundColor Yellow

                            # Build chunking parameters
                            # Use the original source file for chunking, not the converted output
                            $chunkParams = @{
                                InputPath = $item.FilePath  # Use original document
                                TokenizerBackend = $item.ChunkTokenizerBackend
                                MaxTokens = $item.ChunkMaxTokens
                                MergePeers = $item.ChunkMergePeers
                                TableSerialization = $item.ChunkTableSerialization
                                PictureStrategy = $item.ChunkPictureStrategy
                            }

                            if ($item.ChunkTokenizerBackend -eq 'hf') {
                                $chunkParams.TokenizerModel = $item.ChunkTokenizerModel
                            } else {
                                $chunkParams.OpenAIModel = $item.ChunkOpenAIModel
                            }

                            if ($item.ChunkIncludeContext) {
                                $chunkParams.IncludeContext = $true
                            }

                            # Generate chunks output path in the same directory as the converted file
                            $baseNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($outputFile)
                            $outputDir = [System.IO.Path]::GetDirectoryName($outputFile)
                            $chunksOutputPath = [System.IO.Path]::Combine($outputDir, "$baseNameWithoutExt.chunks.jsonl")
                            $chunkParams.OutputPath = $chunksOutputPath

                            # Invoke chunking
                            $chunkResult = Invoke-DoclingHybridChunking @chunkParams

                            if ($chunkResult -and (Test-Path $chunkResult)) {
                                $statusUpdate.ChunksFile = $chunkResult
                                Write-Host "Chunking completed: $chunkResult" -ForegroundColor Green
                            } else {
                                Write-Warning "Chunking did not produce output file"
                            }
                        }
                        catch {
                            Write-Warning "Chunking failed: $($_.Exception.Message)"
                            # Don't fail the whole process if chunking fails
                            $statusUpdate.ChunkingError = $_.Exception.Message
                        }
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
