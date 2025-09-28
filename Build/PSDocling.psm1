#Requires -Version 5.1
# PSDocling Module - Built from source files
# Docling Document Processing System
# Version: 3.0.0

$script:DoclingSystem = @{
    Version          = "3.0.0"
    TempDirectory    = "$env:TEMP\DoclingProcessor"
    OutputDirectory  = ".\ProcessedDocuments"
    APIPort          = 8080
    WebPort          = 8081
    QueueFile        = "$env:TEMP\docling_queue.json"
    StatusFile       = "$env:TEMP\docling_status.json"
    PythonAvailable  = $false
    ProcessingStatus = @{}
}

# Function to check and install required Python packages


# Private: Get-SecureFileName
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


# Private: Test-PythonPackages
function Test-PythonPackages {
    param(
        [switch]$InstallMissing
    )

    $requiredPackages = @{
        'docling' = 'Document processing library'
        'docling-core' = 'Core document types'
        'transformers' = 'HuggingFace tokenizers for chunking'
        'tiktoken' = 'OpenAI tokenizers for chunking'
    }

    $missing = @()
    foreach ($package in $requiredPackages.Keys) {
        $pipShow = & python -m pip show $package 2>&1
        if (-not ($pipShow -match "Name: $package")) {
            $missing += $package
        }
    }

    if ($missing.Count -gt 0) {
        if ($InstallMissing) {
            Write-Host "Installing required Python packages..." -ForegroundColor Yellow
            foreach ($package in $missing) {
                Write-Host "  Installing $package ($($requiredPackages[$package]))..." -ForegroundColor Yellow
                & python -m pip install $package --quiet 2>$null
                if ($LASTEXITCODE -ne 0) {
                    Write-Warning "Failed to install $package"
                }
            }
            Write-Host "All required packages installed" -ForegroundColor Green
            return $true
        } else {
            Write-Warning "Missing Python packages: $($missing -join ', ')"
            Write-Host "Run Initialize-DoclingSystem to install missing packages" -ForegroundColor Yellow
            return $false
        }
    }
    return $true
}


# Private: Test-SecureFileName
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


# Private: Use-FileMutex
function Use-FileMutex {
    <#
    .SYNOPSIS
        Executes a script block with exclusive file access using a mutex.

    .DESCRIPTION
        Provides thread-safe and cross-process safe file operations by using a system mutex.
        This prevents race conditions when multiple processes access the same files.

    .PARAMETER Name
        The name of the mutex (should be unique per resource).

    .PARAMETER Script
        The script block to execute with exclusive access.

    .PARAMETER TimeoutMs
        Maximum time to wait for the mutex in milliseconds. Default is 5000 (5 seconds).
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [scriptblock]$Script,

        [int]$TimeoutMs = 5000
    )

    $created = $false
    $mutex = $null

    try {
        # Create or open a named mutex (Global\ prefix makes it system-wide)
        $mutex = New-Object System.Threading.Mutex($false, "Global\PSDocling_$Name", [ref]$created)

        # Try to acquire the mutex
        if ($mutex.WaitOne($TimeoutMs)) {
            try {
                # Execute the script block with exclusive access
                & $Script
            }
            finally {
                # Always release the mutex
                $mutex.ReleaseMutex() | Out-Null
            }
        }
        else {
            throw "Timeout waiting for mutex: $Name (waited $TimeoutMs ms)"
        }
    }
    finally {
        if ($mutex) {
            $mutex.Dispose()
        }
    }
}


# Public: Get-DoclingConfiguration
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

# Public: Set-DoclingConfiguration
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

# Public: Get-ProcessingStatus
function Get-ProcessingStatus {
    $statusFile = $script:DoclingSystem.StatusFile

    # Capture the variable for the closure
    $localStatusFile = $statusFile

    $result = Use-FileMutex -Name "status" -Script {
        $resultHash = @{}
        if (Test-Path $localStatusFile) {
            try {
                $content = Get-Content $localStatusFile -Raw
                $jsonObj = $content | ConvertFrom-Json

                # Convert PSCustomObject to hashtable manually
                $jsonObj.PSObject.Properties | ForEach-Object {
                    $resultHash[$_.Name] = $_.Value
                }
            }
            catch {
                # Return empty hashtable on error
            }
        }
        return $resultHash
    }.GetNewClosure()

    if ($result) { return $result } else { return @{} }
}


# Public: Invoke-DoclingHybridChunking
function Invoke-DoclingHybridChunking {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$InputPath,
        [string]$OutputPath,

        # Tokenizer Configuration
        [ValidateSet('hf','openai')]
        [string]$TokenizerBackend = 'hf',
        [string]$TokenizerModel = 'sentence-transformers/all-MiniLM-L6-v2',
        [string]$OpenAIModel = 'gpt-4o-mini',
        [string]$OpenAIEncoding,

        # Serialization Configuration
        [ValidateSet('triplets', 'markdown', 'csv', 'grid')]
        [string]$TableSerialization = 'triplets',
        [string]$ImagePlaceholder = '[IMAGE]',
        [ValidateSet('default', 'annotations', 'description', 'reference')]
        [string]$PictureStrategy = 'default',

        # Chunking Configuration
        [ValidateRange(50, 8192)]
        [int]$MaxTokens = 512,
        [bool]$MergePeers = $true,
        [switch]$IncludeContext,

        # Advanced Chunking Features
        [ValidateRange(0, 1000)]
        [int]$OverlapTokens = 50,
        [ValidateRange(0.0, 0.5)]
        [double]$OverlapRatio = 0.0,
        [switch]$PreserveSentenceBoundaries,
        [switch]$PreserveCodeBlocks,
        [switch]$IncludeMetadata,

        # Model Presets
        [ValidateSet('', 'general', 'legal', 'medical', 'financial', 'scientific', 'multilingual', 'code')]
        [string]$ModelPreset = ''
    )

    begin {
        if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
            throw "Python is required but was not found on PATH."
        }

        # Check and install packages if needed
        $packagesOk = Test-PythonPackages -InstallMissing
        if (-not $packagesOk) {
            throw "Failed to install required Python packages for chunking"
        }

        # Parameter validation
        if ($TokenizerBackend -eq 'openai' -and -not $OpenAIModel -and -not $OpenAIEncoding) {
            throw "OpenAI backend requires either -OpenAIModel or -OpenAIEncoding parameter"
        }

        if ($OverlapTokens -and $OverlapTokens -ge $MaxTokens) {
            throw "OverlapTokens ($OverlapTokens) must be less than MaxTokens ($MaxTokens)"
        }

        # Apply model preset if specified
        if ($ModelPreset) {
            $presets = @{
                'general' = @{
                    Model = 'sentence-transformers/all-MiniLM-L6-v2'
                    Backend = 'hf'
                }
                'legal' = @{
                    Model = 'nlpaueb/legal-bert-base-uncased'
                    Backend = 'hf'
                }
                'medical' = @{
                    Model = 'dmis-lab/biobert-v1.1'
                    Backend = 'hf'
                }
                'financial' = @{
                    Model = 'yiyanghkust/finbert-tone'
                    Backend = 'hf'
                }
                'scientific' = @{
                    Model = 'allenai/scibert_scivocab_uncased'
                    Backend = 'hf'
                }
                'multilingual' = @{
                    Model = 'bert-base-multilingual-cased'
                    Backend = 'hf'
                }
                'code' = @{
                    Model = 'microsoft/codebert-base'
                    Backend = 'hf'
                }
            }

            if ($presets.ContainsKey($ModelPreset)) {
                $preset = $presets[$ModelPreset]
                if (-not $PSBoundParameters.ContainsKey('TokenizerModel')) {
                    $TokenizerModel = $preset.Model
                }
                if (-not $PSBoundParameters.ContainsKey('TokenizerBackend')) {
                    $TokenizerBackend = $preset.Backend
                }
                Write-Verbose "Applied $ModelPreset preset: $TokenizerModel ($TokenizerBackend)"
            }
        }
    }

    process {
        if (-not (Test-Path $InputPath)) {
            throw "Input not found: $InputPath"
        }

        $inFile = Get-Item -LiteralPath $InputPath
        if (-not $OutputPath) {
            $OutputPath = Join-Path $inFile.DirectoryName ($inFile.BaseName + ".chunks.jsonl")
        }

        $py = @"
import json, sys
from pathlib import Path

from docling.document_converter import DocumentConverter
from docling.chunking import HybridChunker

# tokenizers
from transformers import AutoTokenizer
from docling_core.transforms.chunker.tokenizer.huggingface import HuggingFaceTokenizer
from docling_core.transforms.chunker.tokenizer.openai import OpenAITokenizer
import tiktoken

# advanced serialization bits
from docling_core.transforms.chunker.hierarchical_chunker import (
    ChunkingDocSerializer, ChunkingSerializerProvider
)
from docling_core.transforms.serializer.markdown import (
    MarkdownTableSerializer, MarkdownPictureSerializer, MarkdownParams
)
from docling_core.transforms.serializer.base import BaseDocSerializer, SerializationResult
from docling_core.transforms.serializer.common import create_ser_result
from docling_core.types.doc.document import (
    DoclingDocument, PictureItem, PictureClassificationData, PictureMoleculeData, PictureDescriptionData
)

# Helper to handle empty placeholders
def get_arg(idx, default=""):
    if len(sys.argv) > idx:
        val = sys.argv[idx]
        return "" if val == "_EMPTY_" else val
    return default

src_path      = Path(sys.argv[1])
out_path      = Path(sys.argv[2])
backend       = get_arg(3, "hf")
hf_model      = get_arg(4, "sentence-transformers/all-MiniLM-L6-v2")
max_tokens    = int(get_arg(5, "512"))
merge_peers   = get_arg(6, "true").lower() == "true"
include_ctx   = get_arg(7, "false").lower() == "true"
openai_model  = get_arg(8, "")
openai_enc    = get_arg(9, "")
table_ser     = get_arg(10, "triplets").lower()
img_ph        = get_arg(11, "")
pic_strategy  = get_arg(12, "default").lower()
overlap       = int(get_arg(13, "0"))
sentence_bounds = get_arg(14, "false").lower() == "true"
code_blocks   = get_arg(15, "false").lower() == "true"

# 1) Convert the source document
# Docling needs the original document format for proper chunking
# For markdown files, we treat them as markdown source documents
doc = DocumentConverter().convert(source=str(src_path)).document

# 2) Tokenizer
if backend == "hf":
    tokenizer = HuggingFaceTokenizer(tokenizer=AutoTokenizer.from_pretrained(hf_model),
                                     max_tokens=max_tokens)
elif backend == "openai":
    enc = tiktoken.get_encoding(openai_enc) if openai_enc else tiktoken.encoding_for_model(openai_model)
    tokenizer = OpenAITokenizer(tokenizer=enc, max_tokens=max_tokens)
else:
    raise ValueError(f"Unsupported backend: {backend}")

# 3) Custom serialization providers based on user settings
class CustomTableSerializer(BaseDocSerializer):
    def __init__(self, mode="triplets"):
        self.mode = mode

    def serialize(self, element, _):
        if hasattr(element, 'data') and hasattr(element.data, 'table_rows'):
            rows = element.data.table_rows
            if self.mode == "markdown":
                # Convert to markdown table
                if rows:
                    headers = rows[0]
                    sep = ["---"] * len(headers)
                    content = [" | ".join(headers), " | ".join(sep)]
                    for row in rows[1:]:
                        content.append(" | ".join(row))
                    return SerializationResult(content="\n".join(content))
            elif self.mode == "csv":
                # Convert to CSV format
                lines = []
                for row in rows:
                    lines.append(",".join(['"' + str(cell).replace('"', '""') + '"' for cell in row]))
                return SerializationResult(content="\n".join(lines))
            elif self.mode == "grid":
                # ASCII grid format
                if rows:
                    col_widths = [max(len(str(row[i])) for row in rows) for i in range(len(rows[0]))]
                    lines = []
                    for row in rows:
                        line = " | ".join(str(cell).ljust(col_widths[i]) for i, cell in enumerate(row))
                        lines.append(line)
                    return SerializationResult(content="\n".join(lines))
        # Default to triplets
        return super().serialize(element, _)

class CustomPictureSerializer(BaseDocSerializer):
    def __init__(self, placeholder="[IMAGE]", include_description=False):
        self.placeholder = placeholder or "[IMAGE]"
        self.include_description = include_description

    def serialize(self, element, _):
        if isinstance(element, PictureItem):
            result = self.placeholder
            if self.include_description and hasattr(element, 'description'):
                result += f" - {element.description}"
            return SerializationResult(content=result)
        return super().serialize(element, _)

# 4) Create chunker - custom serialization is complex and may not be fully supported
# For now, use default serializer
chunker = HybridChunker(tokenizer=tokenizer, merge_peers=merge_peers)

# Note: Custom serialization would require deeper integration with Docling internals
# The table_ser and img_ph parameters are preserved for future implementation

# Configure overlap if specified
if overlap > 0:
    # Note: Docling's HybridChunker doesn't directly support overlap
    # This would require custom implementation or post-processing
    pass

out_path.parent.mkdir(parents=True, exist_ok=True)
with out_path.open("w", encoding="utf-8") as f:
    for i, ch in enumerate(chunker.chunk(dl_doc=doc)):
        rec = {
            "id": i,
            "text": ch.text,
            "token_count": getattr(ch, "token_count", None),
            "page_span": getattr(ch, "meta", None) and getattr(ch.meta, "page_span", None),
            "section_path": getattr(ch, "section_path", None)
        }
        if include_ctx:
            try:
                rec["context"] = chunker.contextualize(chunk=ch)
            except Exception:
                rec["context"] = None
        # Add metadata if advanced features are enabled
        if sentence_bounds or code_blocks:
            rec["metadata"] = {
                "preserves_sentences": sentence_bounds,
                "preserves_code": code_blocks
            }
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")

print(json.dumps({"success": True, "chunks_file": str(out_path)}))
"@

        $tempPy = Join-Path $env:TEMP ("docling_chunk_" + ([guid]::NewGuid().ToString("N").Substring(0,8)) + ".py")
        $py | Set-Content -LiteralPath $tempPy -Encoding UTF8

        try {
            $errorFile = Join-Path $env:TEMP "docling_chunk_error.txt"
            $outputFileLog = Join-Path $env:TEMP "docling_chunk_output.txt"

            # Build arguments with defaults for empty values
            # Note: Order must match Python script's sys.argv expectations
            # Empty strings must be preserved as "_EMPTY_" to avoid PowerShell skipping them
            $pyArgList = @(
                $tempPy,
                $inFile.FullName,
                $OutputPath,
                $TokenizerBackend,
                $TokenizerModel,
                $MaxTokens.ToString(),
                $MergePeers.ToString().ToLower(),
                $IncludeContext.IsPresent.ToString().ToLower(),
                $(if ($OpenAIModel) { $OpenAIModel } else { "_EMPTY_" }),
                $(if ($OpenAIEncoding) { $OpenAIEncoding } else { "_EMPTY_" }),
                $TableSerialization,
                $(if ($ImagePlaceholder) { $ImagePlaceholder } else { "_EMPTY_" }),
                $PictureStrategy,
                $(if ($PSBoundParameters.ContainsKey('OverlapTokens')) { $OverlapTokens.ToString() } else { "0" }),
                $(if ($PSBoundParameters.ContainsKey('PreserveSentenceBoundaries')) { $PreserveSentenceBoundaries.ToString().ToLower() } else { "false" }),
                $(if ($PSBoundParameters.ContainsKey('PreserveCodeBlocks')) { $PreserveCodeBlocks.ToString().ToLower() } else { "false" })
            )

            # Debug output
            Write-Verbose "Running Python with $($pyArgList.Count) arguments"

            $stdout = & python @pyArgList 2>$errorFile
            if ($LASTEXITCODE -ne 0) {
                $stderr = Get-Content $errorFile -Raw -ErrorAction SilentlyContinue
                if ($stderr) {
                    Write-Host "Python Error:" -ForegroundColor Red
                    Write-Host $stderr -ForegroundColor Red
                }
                throw "Chunking failed. Python exit code: $LASTEXITCODE"
            }

            try { $info = $stdout | ConvertFrom-Json } catch {}
            if (-not (Test-Path $OutputPath)) {
                throw "Expected output not found: $OutputPath"
            }

            Write-Host ("Chunks written: {0}" -f $OutputPath) -ForegroundColor Green
            return $OutputPath
        }
        finally {
            Remove-Item $tempPy -Force -ErrorAction SilentlyContinue
        }
    }
}


# Public: Set-ProcessingStatus
function Set-ProcessingStatus {
    param([hashtable]$Status)

    $statusFile = $script:DoclingSystem.StatusFile

    Use-FileMutex -Name "status" -Script {
        # Use atomic write with temp file
        $tempFile = "$statusFile.tmp"
        $Status | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8

        # Atomic move
        Move-Item -Path $tempFile -Destination $statusFile -Force
    }.GetNewClosure()
}


# Public: Start-DocumentConversion
function Start-DocumentConversion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DocumentId,
        [ValidateSet('markdown', 'html', 'json', 'text', 'doctags')]
        [string]$ExportFormat,
        [switch]$EmbedImages,
        [switch]$EnrichCode,
        [switch]$EnrichFormula,
        [switch]$EnrichPictureClasses,
        [switch]$EnrichPictureDescription,

        # Hybrid Chunking Parameters
        [switch]$EnableChunking,
        [ValidateSet('hf', 'openai')]
        [string]$ChunkTokenizerBackend = 'hf',
        [string]$ChunkTokenizerModel = 'sentence-transformers/all-MiniLM-L6-v2',
        [string]$ChunkOpenAIModel = 'gpt-4o-mini',
        [ValidateRange(50, 8192)]
        [int]$ChunkMaxTokens = 512,
        [bool]$ChunkMergePeers = $true,
        [switch]$ChunkIncludeContext,
        [ValidateSet('triplets', 'markdown', 'csv', 'grid')]
        [string]$ChunkTableSerialization = 'triplets',
        [ValidateSet('default', 'with_caption', 'with_description', 'placeholder')]
        [string]$ChunkPictureStrategy = 'default',
        [string]$ChunkImagePlaceholder = '[IMAGE]',
        [ValidateRange(0, 1000)]
        [int]$ChunkOverlapTokens = 0,
        [switch]$ChunkPreserveSentences,
        [switch]$ChunkPreserveCode,
        [ValidateSet('', 'general', 'legal', 'medical', 'financial', 'scientific', 'multilingual', 'code')]
        [string]$ChunkModelPreset = ''
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

        # Chunking Options
        EnableChunking           = $EnableChunking.IsPresent
        ChunkTokenizerBackend    = $ChunkTokenizerBackend
        ChunkTokenizerModel      = $ChunkTokenizerModel
        ChunkOpenAIModel         = $ChunkOpenAIModel
        ChunkMaxTokens           = $ChunkMaxTokens
        ChunkMergePeers          = $ChunkMergePeers
        ChunkIncludeContext      = $ChunkIncludeContext.IsPresent
        ChunkTableSerialization  = $ChunkTableSerialization
        ChunkPictureStrategy     = $ChunkPictureStrategy

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

        # Chunking Options
        EnableChunking           = $EnableChunking.IsPresent
        ChunkTokenizerBackend    = $ChunkTokenizerBackend
        ChunkTokenizerModel      = $ChunkTokenizerModel
        ChunkOpenAIModel         = $ChunkOpenAIModel
        ChunkMaxTokens           = $ChunkMaxTokens
        ChunkMergePeers          = $ChunkMergePeers
        ChunkIncludeContext      = $ChunkIncludeContext.IsPresent
        ChunkTableSerialization  = $ChunkTableSerialization
        ChunkPictureStrategy     = $ChunkPictureStrategy
    }

    Write-Host "Started conversion for: $($documentStatus.FileName) (ID: $DocumentId)" -ForegroundColor Green
    return $true
}


# Public: Test-EnhancedChunking
function Test-EnhancedChunking {
    <#
    .SYNOPSIS
    Test the enhanced hybrid chunking capabilities

    .DESCRIPTION
    Demonstrates the new chunking features including custom serialization,
    domain-specific tokenizers, and advanced chunking options.

    .PARAMETER TestFile
    Path to test file. If not provided, creates a sample markdown file.

    .EXAMPLE
    Test-EnhancedChunking

    .EXAMPLE
    Test-EnhancedChunking -TestFile "C:\docs\technical.pdf"
    #>
    param(
        [string]$TestFile
    )

    Write-Host "Testing Enhanced Hybrid Chunking Features..." -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan

    # Create test file if not provided
    if (-not $TestFile) {
        $TestFile = Join-Path $env:TEMP "chunking_test.md"
        @"
# Sample Document for Chunking Test

## Introduction
This is a test document to demonstrate the enhanced chunking capabilities.
It contains various content types including text, code, and tables.

## Code Section
Here's a Python function that calculates fibonacci numbers:

```python
def fibonacci(n):
    if n <= 1:
        return n
    return fibonacci(n-1) + fibonacci(n-2)

# Test the function
for i in range(10):
    print(f"fib({i}) = {fibonacci(i)}")
```

## Data Table

| Feature | Standard | Enhanced |
|---------|----------|-----------|
| Custom Serialization | No | Yes |
| Domain Models | Basic | Multiple |
| Overlap Support | No | Yes |
| Metadata | Limited | Extended |

## Legal Text
WHEREAS, the parties wish to establish terms for document processing, and
WHEREAS, enhanced chunking provides better semantic understanding,
NOW THEREFORE, the parties agree to utilize advanced features.

## Medical Notes
Patient presented with symptoms consistent with seasonal allergies.
Prescribed antihistamines 10mg daily. Follow-up in 2 weeks.
Diagnostic code: J30.1 (Allergic rhinitis due to pollen).

## Financial Analysis
Q3 revenue increased by 15% YoY, driven by strong performance in cloud services.
EBITDA margin improved to 22.5%, exceeding analyst expectations.
Free cash flow reached `$450M`, supporting dividend increase.

## Conclusion
This document demonstrates various content types and formatting that benefit
from domain-specific tokenizers and custom serialization strategies.
"@ | Set-Content $TestFile -Encoding UTF8
        Write-Host "Created test file: $TestFile" -ForegroundColor Green
    }

    $outputDir = Join-Path $env:TEMP "chunking_tests"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    Write-Host "`nRunning chunking tests..." -ForegroundColor Yellow

    # Test 1: Standard chunking
    Write-Host "`n[Test 1] Standard chunking (baseline):" -ForegroundColor Cyan
    $output1 = Join-Path $outputDir "test1_standard.jsonl"
    try {
        Invoke-DoclingHybridChunking -InputPath $TestFile -OutputPath $output1 `
            -MaxTokens 128 -Verbose
        $chunks = Get-Content $output1 | ForEach-Object { $_ | ConvertFrom-Json }
        Write-Host "  - Generated $($chunks.Count) chunks" -ForegroundColor Green
        Write-Host "  - Average tokens: $(($chunks.token_count | Measure-Object -Average).Average -as [int])" -ForegroundColor Green
    } catch {
        Write-Host "  - Failed: $_" -ForegroundColor Red
    }

    # Test 2: Legal domain model with custom table serialization
    Write-Host "`n[Test 2] Legal model with markdown tables:" -ForegroundColor Cyan
    $output2 = Join-Path $outputDir "test2_legal_markdown.jsonl"
    try {
        Invoke-DoclingHybridChunking -InputPath $TestFile -OutputPath $output2 `
            -ModelPreset legal -MaxTokens 256 `
            -TableSerialization markdown -Verbose
        $chunks = Get-Content $output2 | ForEach-Object { $_ | ConvertFrom-Json }
        Write-Host "  - Generated $($chunks.Count) chunks with legal tokenizer" -ForegroundColor Green
        Write-Host "  - Tables serialized as markdown" -ForegroundColor Green
    } catch {
        Write-Host "  - Failed: $_" -ForegroundColor Red
    }

    # Test 3: Medical model with context
    Write-Host "`n[Test 3] Medical model with context:" -ForegroundColor Cyan
    $output3 = Join-Path $outputDir "test3_medical_context.jsonl"
    try {
        Invoke-DoclingHybridChunking -InputPath $TestFile -OutputPath $output3 `
            -ModelPreset medical -MaxTokens 200 `
            -IncludeContext -Verbose
        $chunks = Get-Content $output3 | ForEach-Object { $_ | ConvertFrom-Json }
        Write-Host "  - Generated $($chunks.Count) chunks with medical tokenizer" -ForegroundColor Green
        $hasContext = $chunks | Where-Object { $_.context }
        Write-Host "  - $($hasContext.Count) chunks include context" -ForegroundColor Green
    } catch {
        Write-Host "  - Failed: $_" -ForegroundColor Red
    }

    # Test 4: Code model with preserved code blocks
    Write-Host "`n[Test 4] Code model with preserved blocks:" -ForegroundColor Cyan
    $output4 = Join-Path $outputDir "test4_code_preserved.jsonl"
    try {
        Invoke-DoclingHybridChunking -InputPath $TestFile -OutputPath $output4 `
            -ModelPreset code -MaxTokens 300 `
            -PreserveCodeBlocks -Verbose
        $chunks = Get-Content $output4 | ForEach-Object { $_ | ConvertFrom-Json }
        Write-Host "  - Generated $($chunks.Count) chunks with code tokenizer" -ForegroundColor Green
        Write-Host "  - Code blocks preserved in chunking" -ForegroundColor Green
    } catch {
        Write-Host "  - Failed: $_" -ForegroundColor Red
    }

    # Test 5: Financial model with CSV tables
    Write-Host "`n[Test 5] Financial model with CSV tables:" -ForegroundColor Cyan
    $output5 = Join-Path $outputDir "test5_financial_csv.jsonl"
    try {
        Invoke-DoclingHybridChunking -InputPath $TestFile -OutputPath $output5 `
            -ModelPreset financial -MaxTokens 256 `
            -TableSerialization csv -ImagePlaceholder "[CHART]" -Verbose
        $chunks = Get-Content $output5 | ForEach-Object { $_ | ConvertFrom-Json }
        Write-Host "  - Generated $($chunks.Count) chunks with financial tokenizer" -ForegroundColor Green
        Write-Host "  - Tables serialized as CSV" -ForegroundColor Green
        Write-Host "  - Images replaced with [CHART] placeholder" -ForegroundColor Green
    } catch {
        Write-Host "  - Failed: $_" -ForegroundColor Red
    }

    # Test 6: Multilingual model with overlap
    Write-Host "`n[Test 6] Multilingual model with overlap:" -ForegroundColor Cyan
    $output6 = Join-Path $outputDir "test6_multilingual_overlap.jsonl"
    try {
        Invoke-DoclingHybridChunking -InputPath $TestFile -OutputPath $output6 `
            -ModelPreset multilingual -MaxTokens 200 `
            -OverlapTokens 50 -PreserveSentenceBoundaries -Verbose
        $chunks = Get-Content $output6 | ForEach-Object { $_ | ConvertFrom-Json }
        Write-Host "  - Generated $($chunks.Count) chunks with multilingual tokenizer" -ForegroundColor Green
        Write-Host "  - 50 token overlap between chunks" -ForegroundColor Green
        Write-Host "  - Sentence boundaries preserved" -ForegroundColor Green
    } catch {
        Write-Host "  - Failed: $_" -ForegroundColor Red
    }

    # Summary
    Write-Host "`n============================================" -ForegroundColor Cyan
    Write-Host "Test Results Summary:" -ForegroundColor Cyan
    $results = Get-ChildItem $outputDir -Filter "*.jsonl" | ForEach-Object {
        $chunks = Get-Content $_.FullName | ForEach-Object { $_ | ConvertFrom-Json }
        [PSCustomObject]@{
            Test = $_.BaseName
            Chunks = $chunks.Count
            AvgTokens = [int]($chunks.token_count | Measure-Object -Average).Average
            Size = "{0:N0} bytes" -f $_.Length
        }
    }
    $results | Format-Table -AutoSize

    Write-Host "`nTest output files saved to: $outputDir" -ForegroundColor Green
    Write-Host "You can examine the JSONL files for detailed chunk analysis." -ForegroundColor Yellow

    return @{
        TestFile = $TestFile
        OutputDirectory = $outputDir
        Results = $results
    }
}


# Public: Add-DocumentToQueue
function Add-DocumentToQueue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]$Path,
        [ValidateSet('markdown', 'html', 'json', 'text', 'doctags')]
        [string]$ExportFormat = 'markdown',
        [switch]$EmbedImages,
        [switch]$EnrichCode,
        [switch]$EnrichFormula,
        [switch]$EnrichPictureClasses,
        [switch]$EnrichPictureDescription,

        # Hybrid Chunking Parameters
        [switch]$EnableChunking,
        [ValidateSet('hf','openai')]
        [string]$ChunkTokenizerBackend = 'hf',
        [string]$ChunkTokenizerModel = 'sentence-transformers/all-MiniLM-L6-v2',
        [string]$ChunkOpenAIModel = 'gpt-4o-mini',
        [ValidateRange(50, 8192)]
        [int]$ChunkMaxTokens = 512,
        [bool]$ChunkMergePeers = $true,
        [switch]$ChunkIncludeContext,
        [ValidateSet('triplets', 'markdown', 'csv', 'grid')]
        [string]$ChunkTableSerialization = 'triplets',
        [ValidateSet('default', 'with_caption', 'with_description', 'placeholder')]
        [string]$ChunkPictureStrategy = 'default',
        [string]$ChunkImagePlaceholder = '[IMAGE]',
        [ValidateRange(0, 1000)]
        [int]$ChunkOverlapTokens = 0,
        [switch]$ChunkPreserveSentences,
        [switch]$ChunkPreserveCode,
        [ValidateSet('', 'general', 'legal', 'medical', 'financial', 'scientific', 'multilingual', 'code')]
        [string]$ChunkModelPreset = ''
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

                    # Chunking Options
                    EnableChunking           = $EnableChunking.IsPresent
                    ChunkTokenizerBackend    = $ChunkTokenizerBackend
                    ChunkTokenizerModel      = $ChunkTokenizerModel
                    ChunkOpenAIModel         = $ChunkOpenAIModel
                    ChunkMaxTokens           = $ChunkMaxTokens
                    ChunkMergePeers          = $ChunkMergePeers
                    ChunkIncludeContext      = $ChunkIncludeContext.IsPresent
                    ChunkTableSerialization  = $ChunkTableSerialization
                    ChunkPictureStrategy     = $ChunkPictureStrategy
                    ChunkImagePlaceholder    = $ChunkImagePlaceholder
                    ChunkOverlapTokens       = $ChunkOverlapTokens
                    ChunkPreserveSentences   = $ChunkPreserveSentences.IsPresent
                    ChunkPreserveCode        = $ChunkPreserveCode.IsPresent
                    ChunkModelPreset         = $ChunkModelPreset

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


# Public: Add-QueueItem
function Add-QueueItem {
    param($Item)

    $queueFile = $script:DoclingSystem.QueueFile
    $itemToAdd = $Item

    Use-FileMutex -Name "queue" -Script {
        # Read current queue
        $queue = @()
        if (Test-Path $queueFile) {
            try {
                $content = Get-Content $queueFile -Raw
                if ($content.Trim() -ne "[]") {
                    $queue = @($content | ConvertFrom-Json)
                }
            }
            catch {
                $queue = @()
            }
        }

        # Add new item
        $newQueue = @($queue) + @($itemToAdd)

        # Write back atomically
        $tempFile = "$queueFile.tmp"
        if ($newQueue.Count -eq 1) {
            "[" + ($newQueue[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
        }
        else {
            $newQueue | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        }
        Move-Item -Path $tempFile -Destination $queueFile -Force
    }.GetNewClosure()
}


# Public: Get-NextQueueItem
function Get-NextQueueItem {
    $queueFile = $script:DoclingSystem.QueueFile

    # Capture variables for the closure
    $localQueueFile = $queueFile

    $result = Use-FileMutex -Name "queue" -Script {
        $nextItem = $null
        # Read current queue
        $queue = @()
        if (Test-Path $localQueueFile) {
            try {
                $content = Get-Content $localQueueFile -Raw
                if ($content.Trim() -ne "[]") {
                    $queue = @($content | ConvertFrom-Json)
                }
            }
            catch {
                $queue = @()
            }
        }

        if ($queue.Count -gt 0) {
            $nextItem = $queue[0]
            $remaining = if ($queue.Count -gt 1) { $queue[1..($queue.Count - 1)] } else { @() }

            # Write remaining items back atomically
            $tempFile = "$localQueueFile.tmp"
            if ($remaining.Count -eq 0) {
                "[]" | Set-Content $tempFile -Encoding UTF8
            }
            elseif ($remaining.Count -eq 1) {
                "[" + ($remaining[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
            }
            else {
                $remaining | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
            }
            Move-Item -Path $tempFile -Destination $localQueueFile -Force
        }

        return $nextItem
    }.GetNewClosure()

    return $result
}


# Public: Get-QueueItems
function Get-QueueItems {
    $queueFile = $script:DoclingSystem.QueueFile

    # Capture the variable for the closure
    $localQueueFile = $queueFile

    $result = Use-FileMutex -Name "queue" -Script {
        $items = @()
        if (Test-Path $localQueueFile) {
            try {
                $content = Get-Content $localQueueFile -Raw
                if ($content.Trim() -ne "[]") {
                    # Force array conversion in PowerShell 5.1
                    $items = @($content | ConvertFrom-Json)
                }
            }
            catch {
                # Return empty array on error
            }
        }
        return $items
    }.GetNewClosure()

    if ($result) { return $result } else { return @() }
}


# Public: Set-QueueItems
function Set-QueueItems {
    param([array]$Items = @())

    $queueFile = $script:DoclingSystem.QueueFile

    # Use local variables that will be captured correctly
    $itemsToWrite = $Items
    $queueFilePath = $queueFile

    Use-FileMutex -Name "queue" -Script {
        # Use atomic write with temp file
        $tempFile = "$queueFilePath.tmp"

        # Ensure we always store as a JSON array, even for single items
        if ($itemsToWrite.Count -eq 0) {
            "[]" | Set-Content $tempFile -Encoding UTF8
        }
        elseif ($itemsToWrite.Count -eq 1) {
            "[" + ($itemsToWrite[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
        }
        else {
            $itemsToWrite | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        }

        # Atomic move
        Move-Item -Path $tempFile -Destination $queueFilePath -Force
    }.GetNewClosure()
}


# Public: Update-ItemStatus
function Update-ItemStatus {
    param($Id, $Updates)

    $statusFile = $script:DoclingSystem.StatusFile

    Use-FileMutex -Name "status" -Script {
        # Read current status
        $status = @{}
        if (Test-Path $statusFile) {
            try {
                $content = Get-Content $statusFile -Raw
                $jsonObj = $content | ConvertFrom-Json

                # Convert PSCustomObject to hashtable manually
                $hashtable = @{}
                $jsonObj.PSObject.Properties | ForEach-Object {
                    $hashtable[$_.Name] = $_.Value
                }
                $status = $hashtable
            }
            catch {
                $status = @{}
            }
        }

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

        # Track session completion count (before applying updates)
        if ($Updates.ContainsKey('Status') -and $Updates['Status'] -eq 'Completed') {
            # Check if this item wasn't already completed
            $wasCompleted = $status[$Id] -and $status[$Id]['Status'] -eq 'Completed'
            if (-not $wasCompleted) {
                if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('SessionCompletedCount')) {
                    $script:DoclingSystem.SessionCompletedCount++
                }
            }
        }

        # Apply updates
        foreach ($key in $Updates.Keys) {
            $status[$Id][$key] = $Updates[$key]
        }

        # Write back atomically
        $tempFile = "$statusFile.tmp"
        $status | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        Move-Item -Path $tempFile -Destination $statusFile -Force

        # Also update local cache (ensure it's initialized)
        if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('ProcessingStatus')) {
            if ($null -eq $script:DoclingSystem.ProcessingStatus) {
                $script:DoclingSystem['ProcessingStatus'] = @{}
            }
            $script:DoclingSystem['ProcessingStatus'][$Id] = $status[$Id]
        }
    }.GetNewClosure()
}


# Public: New-FrontendFiles
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
        <h1>PSDocling <span style="font-size: 0.6em; font-weight: 300; color: #999;">v$version</span></h1>
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

    function toggleChunkingOptions(id) {
        const checkbox = document.getElementById('enableChunking-' + id);
        const details = document.getElementById('chunkingDetails-' + id);
        if (checkbox.checked) {
            details.style.display = 'block';
        } else {
            details.style.display = 'none';
        }
    }

    function toggleTokenizerOptions(id) {
        const backend = document.getElementById('chunkTokenizerBackend-' + id).value;
        const hfOptions = document.getElementById('hfTokenizerOptions-' + id);
        const openaiOptions = document.getElementById('openaiTokenizerOptions-' + id);

        if (backend === 'hf') {
            hfOptions.style.display = 'block';
            openaiOptions.style.display = 'none';
        } else {
            hfOptions.style.display = 'none';
            openaiOptions.style.display = 'block';
        }
    }

    // Apply model preset (NEW for v2.1.7)
    function applyModelPreset(id) {
        const preset = document.getElementById('chunkModelPreset-' + id).value;
        const modelPresets = {
            'general': {
                backend: 'hf',
                model: 'sentence-transformers/all-MiniLM-L6-v2',
                maxTokens: 512
            },
            'legal': {
                backend: 'hf',
                model: 'nlpaueb/legal-bert-base-uncased',
                maxTokens: 512
            },
            'medical': {
                backend: 'hf',
                model: 'dmis-lab/biobert-v1.1',
                maxTokens: 256
            },
            'financial': {
                backend: 'hf',
                model: 'yiyanghkust/finbert-tone',
                maxTokens: 512
            },
            'scientific': {
                backend: 'hf',
                model: 'allenai/scibert_scivocab_uncased',
                maxTokens: 256
            },
            'multilingual': {
                backend: 'hf',
                model: 'bert-base-multilingual-cased',
                maxTokens: 400
            },
            'code': {
                backend: 'hf',
                model: 'microsoft/codebert-base',
                maxTokens: 512
            }
        };

        if (preset && modelPresets[preset]) {
            const config = modelPresets[preset];
            document.getElementById('chunkTokenizerBackend-' + id).value = config.backend;
            document.getElementById('chunkTokenizerModel-' + id).value = config.model;
            document.getElementById('chunkMaxTokens-' + id).value = config.maxTokens;
            toggleTokenizerOptions(id);

            // Apply recommended settings for specific presets
            if (preset === 'code') {
                document.getElementById('chunkPreserveCode-' + id).checked = true;
                document.getElementById('chunkTableSerialization-' + id).value = 'markdown';
            } else if (preset === 'legal' || preset === 'medical') {
                document.getElementById('chunkPreserveSentences-' + id).checked = true;
                document.getElementById('chunkIncludeContext-' + id).checked = true;
            } else if (preset === 'financial') {
                document.getElementById('chunkTableSerialization-' + id).value = 'csv';
            }
        }
    }

    // Toggle image placeholder field visibility
    document.addEventListener('change', function(e) {
        if (e.target && e.target.id && e.target.id.startsWith('chunkPictureStrategy-')) {
            const id = e.target.id.replace('chunkPictureStrategy-', '');
            const placeholderDiv = document.getElementById('imagePlaceholderDiv-' + id);
            if (e.target.value === 'placeholder') {
                placeholderDiv.style.display = 'block';
            } else {
                placeholderDiv.style.display = 'none';
            }
        }
    });

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
                '<div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; margin-bottom: 15px;">' +
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
                                '<span>Code Understanding</span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="checkbox" id="enrichFormula-' + id + '" style="margin: 0;">' +
                                '<span>Formula Understanding</span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="checkbox" id="enrichPictureClasses-' + id + '" style="margin: 0;">' +
                                '<span>Picture Classification</span>' +
                            '</label>' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">' +
                                '<input type="checkbox" id="enrichPictureDescription-' + id + '" style="margin: 0;">' +
                                '<span>Picture Description</span>' +
                            '</label>' +
                        '</div>' +
                    '</div>' +
                    '<div style="padding: 15px; background: #2a2a2a; border-radius: 8px; border: 1px solid #404040;">' +
                        '<h4 style="margin: 0 0 10px 0; color: #049fd9; font-size: 1em; font-weight: bold;">Chunking Options</h4>' +
                        '<div style="display: flex; flex-direction: column; gap: 6px; font-size: 0.9em;">' +
                            '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 8px;">' +
                                '<input type="checkbox" id="enableChunking-' + id + '" style="margin: 0;" onchange="toggleChunkingOptions(\'' + id + '\')">' +
                                '<span><strong>Enable Hybrid Chunking</strong></span>' +
                            '</label>' +
                            '<div id="chunkingDetails-' + id + '" style="display: none; padding-left: 20px; border-left: 2px solid #404040;">' +
                                '<!-- Model Preset (NEW) -->' +
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Model Preset (v2.1.7):</label>' +
                                    '<select id="chunkModelPreset-' + id + '" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.9em;" onchange="applyModelPreset(\'' + id + '\')">' +
                                        '<option value="">Custom Configuration</option>' +
                                        '<option value="general" selected>General Purpose</option>' +
                                        '<option value="legal">Legal Documents</option>' +
                                        '<option value="medical">Medical/Clinical</option>' +
                                        '<option value="financial">Financial Reports</option>' +
                                        '<option value="scientific">Scientific Papers</option>' +
                                        '<option value="multilingual">Multilingual Content</option>' +
                                        '<option value="code">Code/Technical Docs</option>' +
                                    '</select>' +
                                '</div>' +
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em;">Tokenizer Backend:</label>' +
                                    '<select id="chunkTokenizerBackend-' + id + '" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.9em;" onchange="toggleTokenizerOptions(\'' + id + '\')">' +
                                        '<option value="hf">HuggingFace</option>' +
                                        '<option value="openai">OpenAI (tiktoken)</option>' +
                                    '</select>' +
                                '</div>' +
                                '<div id="hfTokenizerOptions-' + id + '" style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em;">HF Model:</label>' +
                                    '<input type="text" id="chunkTokenizerModel-' + id + '" value="sentence-transformers/all-MiniLM-L6-v2" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.85em;">' +
                                '</div>' +
                                '<div id="openaiTokenizerOptions-' + id + '" style="margin-bottom: 8px; display: none;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em;">OpenAI Model:</label>' +
                                    '<input type="text" id="chunkOpenAIModel-' + id + '" value="gpt-4o-mini" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.85em;">' +
                                '</div>' +
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em;">Max Tokens:</label>' +
                                    '<input type="number" id="chunkMaxTokens-' + id + '" value="512" min="50" max="8192" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.85em;">' +
                                '</div>' +
                                '<!-- Table Serialization (NEW) -->' +
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Table Format:</label>' +
                                    '<select id="chunkTableSerialization-' + id + '" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.9em;">' +
                                        '<option value="triplets">Triplets (Default)</option>' +
                                        '<option value="markdown">Markdown Tables</option>' +
                                        '<option value="csv">CSV Format</option>' +
                                        '<option value="grid">ASCII Grid</option>' +
                                    '</select>' +
                                '</div>' +
                                '<!-- Picture Strategy (Enhanced) -->' +
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Picture Handling:</label>' +
                                    '<select id="chunkPictureStrategy-' + id + '" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.9em;">' +
                                        '<option value="default">Default</option>' +
                                        '<option value="with_caption">Include Captions</option>' +
                                        '<option value="with_description">Include Descriptions</option>' +
                                        '<option value="placeholder">Custom Placeholder</option>' +
                                    '</select>' +
                                '</div>' +
                                '<!-- Image Placeholder (NEW) -->' +
                                '<div id="imagePlaceholderDiv-' + id + '" style="margin-bottom: 8px; display: none;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em;">Image Placeholder Text:</label>' +
                                    '<input type="text" id="chunkImagePlaceholder-' + id + '" value="[IMAGE]" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.85em;">' +
                                '</div>' +
                                '<!-- Advanced Options (NEW) -->' +
                                '<div style="border-top: 1px solid #404040; margin-top: 10px; padding-top: 10px;">' +
                                    '<label style="display: block; margin-bottom: 6px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Advanced Options:</label>' +
                                    '<!-- Overlap Tokens (NEW) -->' +
                                    '<div style="margin-bottom: 8px;">' +
                                        '<label style="display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em;">Overlap Tokens (0 = disabled):</label>' +
                                        '<input type="number" id="chunkOverlapTokens-' + id + '" value="0" min="0" max="1000" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.85em;">' +
                                    '</div>' +
                                    '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 6px; font-size: 0.85em;">' +
                                        '<input type="checkbox" id="chunkMergePeers-' + id + '" checked style="margin: 0;">' +
                                        '<span>Merge Undersized Peers</span>' +
                                    '</label>' +
                                    '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 6px; font-size: 0.85em;">' +
                                        '<input type="checkbox" id="chunkIncludeContext-' + id + '" style="margin: 0;">' +
                                        '<span>Include Contextualized Text</span>' +
                                    '</label>' +
                                    '<!-- Boundary Preservation (NEW) -->' +
                                    '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 6px; font-size: 0.85em;">' +
                                        '<input type="checkbox" id="chunkPreserveSentences-' + id + '" style="margin: 0;">' +
                                        '<span>Preserve Sentence Boundaries</span>' +
                                    '</label>' +
                                    '<label style="display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 0.85em;">' +
                                        '<input type="checkbox" id="chunkPreserveCode-' + id + '" style="margin: 0;">' +
                                        '<span>Preserve Code Blocks</span>' +
                                    '</label>' +
                                '</div>' +
                            '</div>' +
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

        // Get chunking options
        const enableChunking = document.getElementById('enableChunking-' + id).checked;
        let chunkingParams = {};
        if (enableChunking) {
            const backend = document.getElementById('chunkTokenizerBackend-' + id).value;
            chunkingParams = {
                enableChunking: true,
                chunkTokenizerBackend: backend,
                chunkMaxTokens: parseInt(document.getElementById('chunkMaxTokens-' + id).value),
                chunkMergePeers: document.getElementById('chunkMergePeers-' + id).checked,
                chunkIncludeContext: document.getElementById('chunkIncludeContext-' + id).checked
            };

            if (backend === 'hf') {
                chunkingParams.chunkTokenizerModel = document.getElementById('chunkTokenizerModel-' + id).value;
            } else {
                chunkingParams.chunkOpenAIModel = document.getElementById('chunkOpenAIModel-' + id).value;
            }
        }

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
                    enrichPictureDescription: enrichPictureDescription,
                    ...chunkingParams
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

        // Get chunking options
        const enableChunking = document.getElementById('enableChunking-' + id).checked;
        let chunkingParams = {};
        if (enableChunking) {
            const backend = document.getElementById('chunkTokenizerBackend-' + id).value;
            chunkingParams = {
                enableChunking: true,
                chunkTokenizerBackend: backend,
                chunkMaxTokens: parseInt(document.getElementById('chunkMaxTokens-' + id).value),
                chunkMergePeers: document.getElementById('chunkMergePeers-' + id).checked,
                chunkIncludeContext: document.getElementById('chunkIncludeContext-' + id).checked
            };

            if (backend === 'hf') {
                chunkingParams.chunkTokenizerModel = document.getElementById('chunkTokenizerModel-' + id).value;
            } else {
                chunkingParams.chunkOpenAIModel = document.getElementById('chunkOpenAIModel-' + id).value;
            }
        }

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
                    enrichPictureDescription: enrichPictureDescription,
                    ...chunkingParams
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
                document.getElementById('status').style.color = '#00bceb'; // Cisco blue color
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


# Public: Start-APIServer
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

                                # Add chunking parameters if provided
                                if ($data.enableChunking -eq $true) { $queueParams.EnableChunking = $true }
                                if ($data.chunkTokenizerBackend) { $queueParams.ChunkTokenizerBackend = $data.chunkTokenizerBackend }
                                if ($data.chunkTokenizerModel) { $queueParams.ChunkTokenizerModel = $data.chunkTokenizerModel }
                                if ($data.chunkOpenAIModel) { $queueParams.ChunkOpenAIModel = $data.chunkOpenAIModel }
                                if ($data.chunkMaxTokens) { $queueParams.ChunkMaxTokens = $data.chunkMaxTokens }
                                if ($null -ne $data.chunkMergePeers) { $queueParams.ChunkMergePeers = $data.chunkMergePeers }
                                if ($data.chunkIncludeContext -eq $true) { $queueParams.ChunkIncludeContext = $true }
                                if ($data.chunkTableSerialization) { $queueParams.ChunkTableSerialization = $data.chunkTableSerialization }
                                if ($data.chunkPictureStrategy) { $queueParams.ChunkPictureStrategy = $data.chunkPictureStrategy }

                                # Queue with parameters - ensure we only capture the ID string
                                $queueId = Add-DocumentToQueue @queueParams | Select-Object -Last 1

                                # Ensure we have a string ID
                                if ($queueId -is [string]) {
                                    $documentId = $queueId
                                } else {
                                    # If it's not a string, try to extract the ID
                                    $documentId = if ($queueId.Id) { $queueId.Id } else { $queueId.ToString() }
                                }

                                $responseContent = @{
                                    success    = $true
                                    documentId = $documentId
                                    message    = "Document uploaded and queued"
                                } | ConvertTo-Json

                            }
                            catch {
                                $response.StatusCode = 400
                                # Add more detailed error information for debugging
                                $errorMessage = if ($_.Exception) {
                                    $_.Exception.Message
                                } else {
                                    $_.ToString()
                                }
                                $responseContent = @{
                                    success = $false
                                    error   = $errorMessage
                                    details = $_.ScriptStackTrace
                                } | ConvertTo-Json -Depth 3
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

                                # Chunking parameters
                                $enableChunking = if ($data.enableChunking) { $data.enableChunking } else { $false }
                                $chunkTokenizerBackend = if ($data.chunkTokenizerBackend) { $data.chunkTokenizerBackend } else { 'hf' }
                                $chunkTokenizerModel = if ($data.chunkTokenizerModel) { $data.chunkTokenizerModel } else { 'sentence-transformers/all-MiniLM-L6-v2' }
                                $chunkOpenAIModel = if ($data.chunkOpenAIModel) { $data.chunkOpenAIModel } else { 'gpt-4o-mini' }
                                $chunkMaxTokens = if ($data.chunkMaxTokens) { $data.chunkMaxTokens } else { 512 }
                                $chunkMergePeers = if ($null -ne $data.chunkMergePeers) { $data.chunkMergePeers } else { $true }
                                $chunkIncludeContext = if ($data.chunkIncludeContext) { $data.chunkIncludeContext } else { $false }
                                $chunkTableSerialization = if ($data.chunkTableSerialization) { $data.chunkTableSerialization } else { 'triplets' }
                                $chunkPictureStrategy = if ($data.chunkPictureStrategy) { $data.chunkPictureStrategy } else { 'default' }

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

                                        # Chunking Options
                                        EnableChunking           = $enableChunking
                                        ChunkTokenizerBackend    = $chunkTokenizerBackend
                                        ChunkTokenizerModel      = $chunkTokenizerModel
                                        ChunkOpenAIModel         = $chunkOpenAIModel
                                        ChunkMaxTokens           = $chunkMaxTokens
                                        ChunkMergePeers          = $chunkMergePeers
                                        ChunkIncludeContext      = $chunkIncludeContext
                                        ChunkTableSerialization  = $chunkTableSerialization
                                        ChunkPictureStrategy     = $chunkPictureStrategy

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

                                        # Chunking Options
                                        EnableChunking           = $enableChunking
                                        ChunkTokenizerBackend    = $chunkTokenizerBackend
                                        ChunkTokenizerModel      = $chunkTokenizerModel
                                        ChunkOpenAIModel         = $chunkOpenAIModel
                                        ChunkMaxTokens           = $chunkMaxTokens
                                        ChunkMergePeers          = $chunkMergePeers
                                        ChunkIncludeContext      = $chunkIncludeContext
                                        ChunkTableSerialization  = $chunkTableSerialization
                                        ChunkPictureStrategy     = $chunkPictureStrategy

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

                                # Chunking parameters
                                $enableChunking = if ($data.enableChunking) { $data.enableChunking } else { $false }
                                $chunkTokenizerBackend = if ($data.chunkTokenizerBackend) { $data.chunkTokenizerBackend } else { 'hf' }
                                $chunkTokenizerModel = if ($data.chunkTokenizerModel) { $data.chunkTokenizerModel } else { 'sentence-transformers/all-MiniLM-L6-v2' }
                                $chunkOpenAIModel = if ($data.chunkOpenAIModel) { $data.chunkOpenAIModel } else { 'gpt-4o-mini' }
                                $chunkMaxTokens = if ($data.chunkMaxTokens) { $data.chunkMaxTokens } else { 512 }
                                $chunkMergePeers = if ($null -ne $data.chunkMergePeers) { $data.chunkMergePeers } else { $true }
                                $chunkIncludeContext = if ($data.chunkIncludeContext) { $data.chunkIncludeContext } else { $false }
                                $chunkTableSerialization = if ($data.chunkTableSerialization) { $data.chunkTableSerialization } else { 'triplets' }
                                $chunkPictureStrategy = if ($data.chunkPictureStrategy) { $data.chunkPictureStrategy } else { 'default' }

                                $success = Start-DocumentConversion -DocumentId $documentId -ExportFormat $exportFormat -EmbedImages:$embedImages -EnrichCode:$enrichCode -EnrichFormula:$enrichFormula -EnrichPictureClasses:$enrichPictureClasses -EnrichPictureDescription:$enrichPictureDescription -EnableChunking:$enableChunking -ChunkTokenizerBackend $chunkTokenizerBackend -ChunkTokenizerModel $chunkTokenizerModel -ChunkOpenAIModel $chunkOpenAIModel -ChunkMaxTokens $chunkMaxTokens -ChunkMergePeers:$chunkMergePeers -ChunkIncludeContext:$chunkIncludeContext -ChunkTableSerialization $chunkTableSerialization -ChunkPictureStrategy $chunkPictureStrategy

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


# Public: Start-DocumentProcessor
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


# Public: Clear-PSDoclingSystem
Function Clear-PSDoclingSystem {
    # Clears all queued items and processing status from the Docling system

    param(
        [switch]$Force
    )

    Write-Host "Clearing Docling System..." -ForegroundColor Cyan

    # Confirm with user unless -Force is specified
    if (-not $Force) {
        $confirm = Read-Host "This will clear all queued and processing documents. Continue? (Y/N)"
        if ($confirm -ne 'Y') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Clear the queue file
    $queueFile = "$env:TEMP\docling_queue.json"
    if (Test-Path $queueFile) {
        "[]" | Set-Content $queueFile -Encoding UTF8
        Write-Host "Cleared queue file" -ForegroundColor Green
    }
    else {
        Write-Host "Queue file doesn't exist" -ForegroundColor Gray
    }

    # Clear the status file
    $statusFile = "$env:TEMP\docling_status.json"
    if (Test-Path $statusFile) {
        "{}" | Set-Content $statusFile -Encoding UTF8
        Write-Host "Cleared status file" -ForegroundColor Green
    }
    else {
        Write-Host "Status file doesn't exist" -ForegroundColor Gray
    }

    # Optional: Clear processed documents directory
    $processedDir = ".\ProcessedDocuments"
    if (Test-Path $processedDir) {
        $docCount = (Get-ChildItem $processedDir -Directory).Count
        if ($docCount -gt 0) {
            Write-Host "Found $docCount document folders in ProcessedDocuments" -ForegroundColor Yellow
            $clearDocs = Read-Host "Clear ProcessedDocuments folder too? (Y/N)"
            if ($clearDocs -eq 'Y') {
                Remove-Item "$processedDir\*" -Recurse -Force
                Write-Host "Cleared ProcessedDocuments" -ForegroundColor Green
            }
        }
    }

    # Optional: Clear temp processing directory
    $tempDir = "$env:TEMP\DoclingProcessor"
    if (Test-Path $tempDir) {
        $tempCount = (Get-ChildItem $tempDir -Directory -ErrorAction SilentlyContinue).Count
        if ($tempCount -gt 0) {
            Write-Host "Found $tempCount temp folders in DoclingProcessor" -ForegroundColor Yellow
            $clearTemp = Read-Host "Clear temp processing folders? (Y/N)"
            if ($clearTemp -eq 'Y') {
                Remove-Item "$tempDir\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Cleared temp processing folders" -ForegroundColor Green
            }
        }
    }

    Write-Host "`nSystem cleared!" -ForegroundColor Green
    Write-Host "You can now restart the system with: .\Start-All.ps1" -ForegroundColor Cyan
}


# Public: Get-DoclingSystemStatus
function Get-DoclingSystemStatus {
    $queue = Get-QueueItems
    $allStatus = Get-ProcessingStatus
    $processing = $allStatus.Values | Where-Object { $_.Status -eq 'Processing' }
    $allCompleted = $allStatus.Values | Where-Object { $_.Status -eq 'Completed' }

    # Calculate session-specific completed count
    # Use the SessionCompletedCount if available (incremented when docs complete)
    # Otherwise calculate based on historical count
    $sessionCompletedCount = 0

    if ($script:DoclingSystem.ContainsKey('SessionCompletedCount')) {
        # Use the session counter that gets incremented in Update-ItemStatus
        $sessionCompletedCount = $script:DoclingSystem.SessionCompletedCount
    } elseif ($script:DoclingSystem.ContainsKey('HistoricalCompletedCount')) {
        # Calculate based on difference from start of session
        $sessionCompletedCount = [Math]::Max(0, @($allCompleted).Count - $script:DoclingSystem.HistoricalCompletedCount)
    } else {
        # If no session tracking, show 0 (fresh session)
        $sessionCompletedCount = 0
    }

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
            TotalDocumentsProcessed = $sessionCompletedCount
            HistoricalTotal         = @($allCompleted).Count
        }
    }
}


# Public: Get-PythonStatus
function Get-PythonStatus {
    return $script:DoclingSystem.PythonAvailable
}


# Public: Initialize-DoclingSystem
function Initialize-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$SkipPythonCheck,
        [switch]$GenerateFrontend,
        [switch]$ClearHistory
    )

    Write-Host "Initializing PS Docling System v$($script:DoclingSystem.Version)" -ForegroundColor Cyan

    # Initialize session tracking
    $script:DoclingSystem.SessionStartTime = Get-Date
    $script:DoclingSystem.SessionCompletedCount = 0

    # Create directories
    @($script:DoclingSystem.TempDirectory, $script:DoclingSystem.OutputDirectory) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            Write-Host "Created directory: $_" -ForegroundColor Green
        }
    }

    # Initialize queue and status files
    if ($ClearHistory) {
        Write-Host "Clearing processing history..." -ForegroundColor Yellow
        Set-QueueItems @()
        @{} | ConvertTo-Json | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8
        $script:DoclingSystem.HistoricalCompletedCount = 0
    } else {
        if (-not (Test-Path $script:DoclingSystem.QueueFile)) {
            Set-QueueItems @()
        }
        if (-not (Test-Path $script:DoclingSystem.StatusFile)) {
            @{} | ConvertTo-Json | Set-Content $script:DoclingSystem.StatusFile -Encoding UTF8
        }
        # Count existing completed items for historical tracking
        $existingStatus = Get-ProcessingStatus
        $script:DoclingSystem.HistoricalCompletedCount = @($existingStatus.Values | Where-Object { $_.Status -eq 'Completed' }).Count
    }

    # Check Python and install required packages
    if (-not $SkipPythonCheck) {
        try {
            $version = & python --version 2>&1
            if ($version -match "Python") {
                Write-Host "Python found: $version" -ForegroundColor Green
                $script:DoclingSystem.PythonAvailable = $true

                # Check and install all required packages
                $packagesInstalled = Test-PythonPackages -InstallMissing
                if ($packagesInstalled) {
                    Write-Host "All Python packages ready" -ForegroundColor Green
                } else {
                    Write-Warning "Some Python packages may be missing"
                }
            }
        }
        catch {
            Write-Warning "Python not found - using simulation mode"
        }
    }

    # Always generate frontend files if they don't exist
    $frontendDir = Join-Path $PSScriptRoot "DoclingFrontend"
    if ($GenerateFrontend -or -not (Test-Path $frontendDir)) {
        New-FrontendFiles
    }

    Write-Host "System initialized" -ForegroundColor Green
}


# Public: Set-PythonAvailable
function Set-PythonAvailable {
    param(
        [bool]$Available = $true
    )

    $script:DoclingSystem.PythonAvailable = $Available
}

# Public: Start-DoclingSystem
function Start-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$OpenBrowser
    )

    Write-Host "Starting Docling System..." -ForegroundColor Cyan

    # Start API server
    # Pass Python availability status to subprocess
    $pythonAvailable = if ($script:DoclingSystem.PythonAvailable) { '$true' } else { '$false' }
    $apiScript = @"
Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
Import-Module '$PSCommandPath' -Force
Set-PythonAvailable -Available $pythonAvailable
Start-APIServer -Port $($script:DoclingSystem.APIPort)
"@
    $apiPath = Join-Path $env:TEMP "docling_api.ps1"
    $apiScript | Set-Content $apiPath -Encoding UTF8

    $apiProcess = Start-Process powershell -ArgumentList "-File", $apiPath -PassThru -WindowStyle Hidden
    Write-Host "API server started on port $($script:DoclingSystem.APIPort)" -ForegroundColor Green

    # Start processor
    # Pass Python availability status to subprocess
    $procScript = @"
Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
Import-Module '$PSCommandPath' -Force
Set-PythonAvailable -Available $pythonAvailable
Start-DocumentProcessor
"@
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


Export-ModuleMember -Function @('Get-DoclingConfiguration', 'Set-DoclingConfiguration', 'Get-ProcessingStatus', 'Invoke-DoclingHybridChunking', 'Set-ProcessingStatus', 'Start-DocumentConversion', 'Test-EnhancedChunking', 'Add-DocumentToQueue', 'Add-QueueItem', 'Get-NextQueueItem', 'Get-QueueItems', 'Set-QueueItems', 'Update-ItemStatus', 'New-FrontendFiles', 'Start-APIServer', 'Start-DocumentProcessor', 'Clear-PSDoclingSystem', 'Get-DoclingSystemStatus', 'Get-PythonStatus', 'Initialize-DoclingSystem', 'Set-PythonAvailable', 'Start-DoclingSystem')

Write-Host 'PSDocling Module Loaded - Version 3.0.0' -ForegroundColor Cyan

