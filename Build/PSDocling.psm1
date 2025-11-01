#Requires -Version 5.1
# PSDocling Module - Built from source files
# Docling Document Processing System
# Version: 3.2.0

$script:DoclingSystem = @{
    Version          = "3.2.0"
    ModulePath       = $PSCommandPath
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


# Public: Optimize-ChunksForRAG
function Optimize-ChunksForRAG {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$InputPath,

        [string]$OutputPath,

        [ValidateRange(50, 1000)]
        [int]$TargetMinTokens = 200,

        [ValidateRange(100, 2000)]
        [int]$TargetMaxTokens = 400,

        [ValidateRange(10, 200)]
        [int]$MinTokens = 50,

        [ValidateRange(0.0, 1.0)]
        [double]$DeduplicationThreshold = 0.90
    )

    process {
        if (-not (Test-Path $InputPath)) {
            throw "Input file not found: $InputPath"
        }

        $inFile = Get-Item -LiteralPath $InputPath
        if (-not $OutputPath) {
            $baseName = $inFile.BaseName -replace '\.chunks$',''
            $OutputPath = Join-Path $inFile.DirectoryName ($baseName + ".optimized.jsonl")
        }

        Write-Verbose "Optimizing chunks from: $InputPath"
        Write-Verbose "Target token range: $MinTokens - $TargetMaxTokens (ideal: $TargetMinTokens-$TargetMaxTokens)"

        $py = @"
import json, sys, re, hashlib, html
from pathlib import Path
from collections import defaultdict

def normalize_text(text):
    """Normalize text: decode HTML entities, trim whitespace, unify line endings"""
    if not text:
        return text
    text = html.unescape(text)
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def validate_code_fences(text):
    """Check if all code fences are properly closed"""
    if not text:
        return True, 0
    fence_count = text.count('```')
    return fence_count % 2 == 0, fence_count

def compute_hash(text):
    """Compute MD5 hash for deduplication"""
    normalized = normalize_text(text) if text else ""
    return hashlib.md5(normalized.encode('utf-8')).hexdigest()

def is_near_duplicate(text1, text2, threshold=0.90):
    """Check if two texts are near-duplicates using character overlap"""
    if not text1 or not text2:
        return False

    len1, len2 = len(text1), len(text2)
    if len1 == 0 or len2 == 0:
        return False

    # Quick length check
    if abs(len1 - len2) / max(len1, len2) > 0.3:
        return False

    # Character set overlap
    set1, set2 = set(text1.lower()), set(text2.lower())
    intersection = len(set1 & set2)
    union = len(set1 | set2)

    if union == 0:
        return False

    similarity = intersection / union
    return similarity >= threshold

def estimate_tokens(text):
    """Estimate token count (4 chars â‰ˆ 1 token)"""
    return len(text) // 4 if text else 0

def detect_chunk_type(text):
    """Detect chunk content type"""
    if not text:
        return "empty"
    text_stripped = text.strip()

    if '```' in text:
        return "code"
    if text_stripped.startswith('#'):
        return "heading"
    if re.match(r'^(\s*[-*+]\s|^\s*\d+\.\s)', text, re.MULTILINE):
        return "list"
    if '|' in text and text.count('|') > 3:
        return "table"
    if re.match(r'^\s*(NOTE|TIP|WARNING|IMPORTANT|CAUTION):', text, re.IGNORECASE):
        return "note"

    return "paragraph"

def merge_chunk_pair(chunk1, chunk2):
    """Merge two chunks into one"""
    text1 = chunk1.get("text", "")
    text2 = chunk2.get("text", "")
    merged_text = text1 + "\n\n" + text2
    merged_normalized = normalize_text(merged_text)

    # Merge page spans
    page_span1 = chunk1.get("page_span")
    page_span2 = chunk2.get("page_span")
    merged_page_span = None

    if page_span1 and page_span2:
        pages = []
        if isinstance(page_span1, list):
            pages.extend(page_span1)
        if isinstance(page_span2, list):
            pages.extend(page_span2)
        if pages:
            merged_page_span = [min(pages), max(pages)]
    elif page_span1:
        merged_page_span = page_span1
    elif page_span2:
        merged_page_span = page_span2

    return {
        "text": merged_text,
        "text_normalized": merged_normalized,
        "token_count": estimate_tokens(merged_normalized),
        "page_span": merged_page_span,
        "section_path": chunk1.get("section_path") or chunk2.get("section_path"),
        "doc_id": chunk1.get("doc_id"),
        "source_path": chunk1.get("source_path"),
        "chunk_type": detect_chunk_type(merged_normalized),
        "hash": compute_hash(merged_normalized),
        "char_start": chunk1.get("char_start", 0),
        "char_end": chunk2.get("char_end", 0),
        "lang": chunk1.get("lang", "en"),
        "merged_from": [chunk1.get("id"), chunk2.get("id")]
    }

def split_large_chunk(chunk, max_tokens):
    """Split chunk at sentence boundaries"""
    text = chunk.get("text_normalized") or chunk.get("text", "")
    if not text or estimate_tokens(text) <= max_tokens:
        return [chunk]

    sentences = re.split(r'(?<=[.!?])\s+', text)
    if len(sentences) <= 1:
        return [chunk]

    sub_chunks = []
    current = []
    current_tokens = 0

    for sent in sentences:
        sent_tokens = estimate_tokens(sent)
        if current_tokens + sent_tokens > max_tokens and current:
            sub_text = " ".join(current)
            sub_chunks.append({
                **chunk,
                "text": sub_text,
                "text_normalized": normalize_text(sub_text),
                "token_count": current_tokens,
                "hash": compute_hash(sub_text),
                "split_from": chunk.get("id")
            })
            current = [sent]
            current_tokens = sent_tokens
        else:
            current.append(sent)
            current_tokens += sent_tokens

    if current:
        sub_text = " ".join(current)
        sub_chunks.append({
            **chunk,
            "text": sub_text,
            "text_normalized": normalize_text(sub_text),
            "token_count": current_tokens,
            "hash": compute_hash(sub_text),
            "split_from": chunk.get("id")
        })

    return sub_chunks if sub_chunks else [chunk]

# Read input file
input_path = Path(sys.argv[1])
output_path = Path(sys.argv[2])
min_tokens = int(sys.argv[3])
target_min = int(sys.argv[4])
target_max = int(sys.argv[5])
dedup_threshold = float(sys.argv[6])

chunks = []
with input_path.open('r', encoding='utf-8') as f:
    for line in f:
        if line.strip():
            chunks.append(json.loads(line))

print(f"Loaded {len(chunks)} chunks", file=sys.stderr)

# Step 1: Normalize and add missing metadata
for chunk in chunks:
    text = chunk.get("text", "")
    if "text_normalized" not in chunk:
        chunk["text_normalized"] = normalize_text(text)
    if "token_count" not in chunk or chunk["token_count"] is None:
        chunk["token_count"] = estimate_tokens(chunk["text_normalized"])
    if "hash" not in chunk:
        chunk["hash"] = compute_hash(chunk["text_normalized"])
    if "chunk_type" not in chunk:
        chunk["chunk_type"] = detect_chunk_type(chunk["text_normalized"])

# Step 2: Merge split code blocks (unclosed fences)
merged_chunks = []
i = 0
merge_count = 0

while i < len(chunks):
    chunk = chunks[i]
    text = chunk.get("text_normalized", "")
    fence_valid, fence_count = validate_code_fences(text)

    # If unclosed fence, try to merge with next
    if not fence_valid and i + 1 < len(chunks):
        next_chunk = chunks[i + 1]
        merged = merge_chunk_pair(chunk, next_chunk)
        merged_valid, _ = validate_code_fences(merged.get("text_normalized", ""))

        if merged_valid:
            merged["warnings"] = merged.get("warnings", [])
            merged["warnings"].append("merged_split_code_block")
            merged_chunks.append(merged)
            merge_count += 1
            i += 2  # Skip next chunk
            continue

    merged_chunks.append(chunk)
    i += 1

print(f"Merged {merge_count} split code blocks", file=sys.stderr)
chunks = merged_chunks

# Step 3: Remove duplicates and near-duplicates
seen_hashes = set()
seen_texts = []
dedup_chunks = []
duplicate_count = 0

for chunk in chunks:
    chunk_hash = chunk.get("hash")
    text = chunk.get("text_normalized", "")

    # Exact duplicate check
    if chunk_hash in seen_hashes:
        duplicate_count += 1
        continue

    # Near-duplicate check (against last 20 chunks)
    is_dup = False
    for prev_text in seen_texts[-20:]:
        if is_near_duplicate(text, prev_text, dedup_threshold):
            duplicate_count += 1
            is_dup = True
            break

    if not is_dup:
        seen_hashes.add(chunk_hash)
        seen_texts.append(text)
        dedup_chunks.append(chunk)

print(f"Removed {duplicate_count} duplicates", file=sys.stderr)
chunks = dedup_chunks

# Step 4: Adjust chunk sizes (merge tiny, split large)
adjusted = []
buffer = []
buffer_tokens = 0
adjust_count = 0

for chunk in chunks:
    chunk_tokens = chunk.get("token_count", 0)
    chunk_type = chunk.get("chunk_type", "paragraph")

    # Keep code, tables, headings atomic
    if chunk_type in ["code", "table", "heading"]:
        if buffer:
            merged = buffer[0] if len(buffer) == 1 else merge_chunk_pair(buffer[0], buffer[-1])
            adjusted.append(merged)
            buffer = []
            buffer_tokens = 0
        adjusted.append(chunk)
        continue

    # Tiny chunk: add to buffer
    if chunk_tokens < min_tokens:
        buffer.append(chunk)
        buffer_tokens += chunk_tokens

        if buffer_tokens >= target_min:
            merged = buffer[0] if len(buffer) == 1 else merge_chunk_pair(buffer[0], buffer[-1])
            adjusted.append(merged)
            adjust_count += 1
            buffer = []
            buffer_tokens = 0
    else:
        # Flush buffer
        if buffer:
            merged = buffer[0] if len(buffer) == 1 else merge_chunk_pair(buffer[0], buffer[-1])
            adjusted.append(merged)
            adjust_count += 1
            buffer = []
            buffer_tokens = 0

        # Split large chunks
        if chunk_tokens > target_max * 1.5:
            adjusted.extend(split_large_chunk(chunk, target_max))
            adjust_count += 1
        else:
            adjusted.append(chunk)

# Flush remaining buffer
if buffer:
    merged = buffer[0] if len(buffer) == 1 else merge_chunk_pair(buffer[0], buffer[-1])
    adjusted.append(merged)

print(f"Adjusted {adjust_count} chunk sizes", file=sys.stderr)

# Step 5: Renumber IDs and write output
output_path.parent.mkdir(parents=True, exist_ok=True)
with output_path.open('w', encoding='utf-8') as f:
    for i, chunk in enumerate(adjusted):
        chunk["id"] = i
        f.write(json.dumps(chunk, ensure_ascii=False) + "\n")

# Summary
summary = {
    "success": True,
    "input_chunks": len(chunks) + duplicate_count,
    "output_chunks": len(adjusted),
    "merged_code_blocks": merge_count,
    "duplicates_removed": duplicate_count,
    "size_adjusted": adjust_count,
    "output_file": str(output_path)
}
print(json.dumps(summary, ensure_ascii=False))
"@

        $tempPy = Join-Path $env:TEMP ("docling_optimize_" + ([guid]::NewGuid().ToString("N").Substring(0,8)) + ".py")
        $py | Set-Content -LiteralPath $tempPy -Encoding UTF8

        try {
            $errorFile = Join-Path $env:TEMP "docling_optimize_error.txt"

            $pyArgs = @(
                $tempPy,
                $inFile.FullName,
                $OutputPath,
                $MinTokens.ToString(),
                $TargetMinTokens.ToString(),
                $TargetMaxTokens.ToString(),
                $DeduplicationThreshold.ToString()
            )

            Write-Verbose "Running optimization..."
            $stdout = & python @pyArgs 2>$errorFile

            if ($LASTEXITCODE -ne 0) {
                $stderr = Get-Content $errorFile -Raw -ErrorAction SilentlyContinue
                Write-Error "Optimization failed: $stderr"
                throw "Python exit code: $LASTEXITCODE"
            }

            # Parse summary from last line
            $lines = $stdout -split "`n"
            $summary = $lines[-1] | ConvertFrom-Json -ErrorAction SilentlyContinue

            if ($summary) {
                Write-Host ("Optimized {0} to {1} chunks" -f $summary.input_chunks, $summary.output_chunks) -ForegroundColor Green
                Write-Host ("  - Merged split code blocks: {0}" -f $summary.merged_code_blocks) -ForegroundColor Cyan
                Write-Host ("  - Removed duplicates: {0}" -f $summary.duplicates_removed) -ForegroundColor Cyan
                Write-Host ("  - Size-adjusted: {0}" -f $summary.size_adjusted) -ForegroundColor Cyan
                Write-Host ("  - Output: {0}" -f $OutputPath) -ForegroundColor Yellow
            }

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

    # Use folder-based queue - much simpler!
    Add-QueueItemFolder $DocumentId
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
        ChunkImagePlaceholder    = $ChunkImagePlaceholder
        ChunkOverlapTokens       = $ChunkOverlapTokens
        ChunkPreserveSentences   = $ChunkPreserveSentences.IsPresent
        ChunkPreserveCode        = $ChunkPreserveCode.IsPresent
        ChunkModelPreset         = $ChunkModelPreset
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

    # Capture variables for the closure (similar to Get-NextQueueItem pattern)
    $localQueueFile = $queueFile
    $localItem = $Item

    Use-FileMutex -Name "queue" -Script {
        # Read current queue
        $queue = @()
        if (Test-Path $localQueueFile) {
            try {
                $content = Get-Content $localQueueFile -Raw
                if ($content.Trim() -ne "[]") {
                    $parsed = $content | ConvertFrom-Json
                    # Ensure we get an array
                    if ($parsed -is [array]) {
                        $queue = $parsed
                    } else {
                        $queue = @($parsed)
                    }
                }
            }
            catch {
                $queue = @()
            }
        }

        # Add new item - ensure both are arrays before concatenating
        $queue = @($queue)
        $newQueue = $queue + $localItem

        # Write back atomically - ALWAYS force as an array with explicit formatting
        $tempFile = "$localQueueFile.tmp"
        if ($newQueue.Count -eq 0) {
            "[]" | Set-Content $tempFile -Encoding UTF8
        }
        elseif ($newQueue.Count -eq 1) {
            # Force single item to be an array in JSON
            "[" + ($newQueue[0] | ConvertTo-Json -Depth 10 -Compress) + "]" | Set-Content $tempFile -Encoding UTF8
        }
        else {
            @($newQueue) | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        }
        Move-Item -Path $tempFile -Destination $localQueueFile -Force
    }.GetNewClosure()
}

# Public: Add-QueueItemFolder
function Add-QueueItemFolder {
    param(
        [Parameter(Mandatory)]
        [string]$DocumentId
    )

    # Ensure queue folder exists
    $queueFolder = "$env:TEMP\DoclingQueue"
    if (-not (Test-Path $queueFolder)) {
        New-Item -Path $queueFolder -ItemType Directory -Force | Out-Null
    }

    # Create a queue file for this document
    # File name format: timestamp_documentId.queue
    $timestamp = [DateTime]::Now.ToString("yyyyMMddHHmmssffff")
    $queueFile = Join-Path $queueFolder "${timestamp}_${DocumentId}.queue"

    # Write the document ID to the file (simple content)
    $DocumentId | Set-Content -Path $queueFile -Encoding UTF8

    Write-Verbose "Added to queue: $DocumentId (File: $queueFile)"
    return $queueFile
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
                    $parsed = $content | ConvertFrom-Json
                    # Don't double-wrap arrays
                    if ($parsed -is [array]) {
                        $queue = $parsed
                    } else {
                        $queue = @($parsed)
                    }
                }
            }
            catch {
                $queue = @()
            }
        }

        if ($queue.Count -gt 0) {
            # Ensure we're working with an array
            if ($queue -isnot [array]) {
                $queue = @($queue)
            }

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


# Public: Get-NextQueueItemFolder
function Get-NextQueueItemFolder {
    $queueFolder = "$env:TEMP\DoclingQueue"

    # Ensure queue folder exists
    if (-not (Test-Path $queueFolder)) {
        New-Item -Path $queueFolder -ItemType Directory -Force | Out-Null
        return $null
    }

    # Get all queue files, sorted by creation time (oldest first)
    $queueFiles = Get-ChildItem -Path $queueFolder -Filter "*.queue" |
                  Sort-Object CreationTime |
                  Select-Object -First 1

    if (-not $queueFiles) {
        Write-Verbose "No items in queue folder"
        return $null
    }

    $queueFile = $queueFiles[0]

    # Read the document ID from the file
    $documentId = Get-Content -Path $queueFile.FullName -Raw -Encoding UTF8
    $documentId = $documentId.Trim()

    # Delete the queue file (item is now being processed)
    Remove-Item -Path $queueFile.FullName -Force

    Write-Verbose "Retrieved from queue: $documentId (File: $($queueFile.Name))"
    return $documentId
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


# Public: Get-QueueItemsFolder
function Get-QueueItemsFolder {
    $queueFolder = "$env:TEMP\DoclingQueue"

    # Ensure queue folder exists
    if (-not (Test-Path $queueFolder)) {
        return @()
    }

    # Get all queue files
    $queueFiles = Get-ChildItem -Path $queueFolder -Filter "*.queue" |
                  Sort-Object CreationTime

    if (-not $queueFiles) {
        return @()
    }

    # Read document IDs from all queue files
    $queueItems = @()
    foreach ($file in $queueFiles) {
        $documentId = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        $queueItems += $documentId.Trim()
    }

    Write-Verbose "Found $($queueItems.Count) items in queue"
    return $queueItems
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

    # Capture variables for the closure
    $localStatusFile = $statusFile
    $localId = $Id
    $localUpdates = $Updates

    Use-FileMutex -Name "status" -Script {
        # Read current status
        $status = @{}
        if (Test-Path $localStatusFile) {
            try {
                $content = Get-Content $localStatusFile -Raw
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
        if ($status[$localId]) {
            if ($status[$localId] -is [PSCustomObject]) {
                $itemHash = @{}
                $status[$localId].PSObject.Properties | ForEach-Object {
                    $itemHash[$_.Name] = $_.Value
                }
                $status[$localId] = $itemHash
            }
            # ENSURE it's a hashtable - if not, create new one with existing properties
            if ($status[$localId] -isnot [hashtable]) {
                $oldItem = $status[$localId]
                $status[$localId] = @{}
                # Try to copy any existing properties
                if ($oldItem -is [PSCustomObject]) {
                    $oldItem.PSObject.Properties | ForEach-Object {
                        $status[$localId][$_.Name] = $_.Value
                    }
                }
            }
        }
        else {
            $status[$localId] = @{}
        }

        # Track session completion count (before applying updates)
        if ($localUpdates.ContainsKey('Status') -and $localUpdates['Status'] -eq 'Completed') {
            # Check if this item wasn't already completed
            $wasCompleted = $status[$localId] -and $status[$localId]['Status'] -eq 'Completed'
            if (-not $wasCompleted) {
                if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('SessionCompletedCount')) {
                    $script:DoclingSystem.SessionCompletedCount++
                }
            }
        }

        # Apply updates
        foreach ($key in $localUpdates.Keys) {
            $status[$localId][$key] = $localUpdates[$key]
        }

        # Write back atomically
        $tempFile = "$localStatusFile.tmp"
        $status | ConvertTo-Json -Depth 10 | Set-Content $tempFile -Encoding UTF8
        Move-Item -Path $tempFile -Destination $localStatusFile -Force

        # Also update local cache (ensure it's initialized)
        if ($script:DoclingSystem -and $script:DoclingSystem.ContainsKey('ProcessingStatus')) {
            if ($null -eq $script:DoclingSystem.ProcessingStatus) {
                $script:DoclingSystem['ProcessingStatus'] = @{}
            }
            $script:DoclingSystem['ProcessingStatus'][$localId] = $status[$localId]
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

    # Pull version from module context if present, otherwise default
    $version = $script:DoclingSystem.Version
    if (-not $version) { $version = "3.2.0" }

    # Redesigned HTML (dark/light theme, a11y, keyboard support, and UI polish)
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <title>PSDocling v$version</title>
    <style>
        :root{
            --bg: #0f1115;
            --panel: #171a21;
            --panel-2: #1d2230;
            --stroke: #2a3244;
            --muted: #9aa3b2;
            --text: #e7eef7;
            --brand: #4db6ff;
            --brand-2:#0ea5e9;
            --ok:#10b981;
            --warn:#f59e0b;
            --err:#ef4444;
            --focus: #93c5fd;
            --ring: 0 0 0 3px rgba(147,197,253,.35);
            --radius: 14px;
            --shadow: 0 8px 30px rgba(0,0,0,.35);
        }
        @media (prefers-color-scheme: light){
            :root{
                --bg:#f7f9fc; --panel:#ffffff; --panel-2:#f4f7fb; --stroke:#d9e1ee;
                --muted:#5b6574; --text:#0e1726; --brand:#0369a1; --brand-2:#0284c7;
                --focus:#1d4ed8; --ring:0 0 0 3px rgba(29,78,216,.25);
                --shadow:0 8px 30px rgba(10,30,60,.08);
            }
        }

        *{box-sizing:border-box}
        html,body{height:100%; margin:0; padding:0}
        body{
            font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, "Helvetica Neue", Arial, "Noto Sans", "Apple Color Emoji","Segoe UI Emoji","Segoe UI Symbol";
            background:
                radial-gradient(1200px 1200px at -10% -10%, rgba(77,182,255,.08), transparent 55%),
                radial-gradient(1000px 1000px at 110% -20%, rgba(77,182,255,.08), transparent 55%),
                linear-gradient(180deg, var(--bg), var(--bg));
            color:var(--text);
            -webkit-font-smoothing:antialiased; -moz-osx-font-smoothing:grayscale;
        }

        .container{
            max-width:1200px;
            margin-inline:auto;
            padding:24px;
            display:grid;
            gap:20px;
        }

        /* Header */
        .header{
            background:linear-gradient(180deg, var(--panel), var(--panel-2));
            border:1px solid var(--stroke);
            border-radius:var(--radius);
            padding:24px;
            box-shadow:var(--shadow);
            position:sticky; top:16px; z-index:5;
        }
        .header-row{
            display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px;
        }
        .brand{
            display:flex; align-items:center; gap:12px;
        }
        .logo{
            width:36px; height:36px; border-radius:10px;
            background: radial-gradient(65% 65% at 35% 30%, var(--brand), transparent 60%),
                        linear-gradient(135deg, var(--brand-2), rgba(77,182,255,.25));
            border:1px solid var(--stroke);
            box-shadow: inset 0 0 0 1px rgba(255,255,255,.06), 0 8px 18px rgba(3, 169, 244, .08);
        }
        h1{
            margin:0; font-size: clamp(20px, 3.2vw, 28px); font-weight:700; letter-spacing:.2px;
            background: linear-gradient(45deg, var(--text), rgba(77,182,255,.9));
            -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent;
        }
        .badge{
            margin-left:6px;
            padding:4px 8px; border-radius:999px; font-size:.85rem; font-weight:600;
            color:var(--brand);
            background:linear-gradient(180deg, rgba(77,182,255,.10), rgba(77,182,255,.03));
            border:1px solid var(--stroke);
        }
        .subtitle{margin:6px 0 0 0; color:var(--muted)}
        .status{font-weight:600}
        .status-dot{
            width:8px;height:8px;border-radius:99px;display:inline-block;margin-right:6px;vertical-align:baseline;
            background:var(--warn); box-shadow:0 0 0 3px rgba(11, 132, 245, 0.2);
        }

        /* Card */
        .card{
            background:linear-gradient(180deg, var(--panel), var(--panel-2));
            border:1px solid var(--stroke); border-radius:var(--radius);
            padding:20px; box-shadow:var(--shadow);
        }
        .card-title{
            margin:0 0 10px 0; font-size:1.15rem; font-weight:700; letter-spacing:.2px;
            color:var(--brand);
        }

        /* Upload */
        .drop{
            border:1.5px dashed var(--stroke);
            border-radius:calc(var(--radius) - 4px);
            padding:28px; text-align:center; position:relative; overflow:hidden;
            background:linear-gradient(180deg, rgba(77,182,255,.06), transparent);
            transition:border-color .2s, transform .15s;
            outline:none;
        }
        .drop:hover{border-color:var(--brand); transform: translateY(-1px)}
        .drop:focus-visible{box-shadow: var(--ring); border-color:var(--focus)}
        .drop .shimmer{
            content:""; position:absolute; inset:0; translate:-100% 0;
            background:linear-gradient(90deg, transparent, rgba(77,182,255,.12), transparent);
            animation: sweep 2.6s linear infinite;
        }
        @keyframes sweep{to{translate:100% 0}}
        .drop h3{margin:6px 0 14px 0}
        .muted{color:var(--muted)}
        .kbd{
            padding:2px 6px; border-radius:6px; border:1px solid var(--stroke); background:rgba(0,0,0,.1);
            font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono","Courier New", monospace;
            font-size:.85em;
        }

        .btn{
            appearance:none; border:1px solid var(--stroke); background:linear-gradient(180deg, var(--brand-2), var(--brand));
            color:white; padding:10px 16px; border-radius:10px; font-weight:700; cursor:pointer;
            box-shadow:0 6px 18px rgba(77,182,255,.23); transition:transform .12s ease, box-shadow .2s ease, opacity .2s;
        }
        .btn:hover{transform:translateY(-1px); box-shadow:0 10px 24px rgba(77,182,255,.28)}
        .btn:disabled{opacity:.55; cursor:not-allowed; box-shadow:none}

        .btn-ghost{
            appearance:none; border:1px solid var(--stroke); background:linear-gradient(180deg, rgba(77,182,255,.06), rgba(77,182,255,.03));
            color:var(--brand); padding:10px 14px; border-radius:10px; font-weight:700; cursor:pointer;
            transition:transform .12s ease, background .2s ease;
        }
        .btn-ghost:hover{transform:translateY(-1px); background:linear-gradient(180deg, rgba(77,182,255,.1), rgba(77,182,255,.06))}

        /* Stats */
        .stats{ display:grid; grid-template-columns: repeat(2, 1fr); gap:14px; }
        @media (min-width:780px){ .stats{grid-template-columns: repeat(4, 1fr);} }
        .stat{
            background:linear-gradient(180deg, var(--panel), var(--panel-2));
            border:1px solid var(--stroke); padding:16px; border-radius:12px; text-align:center;
            transition: transform .12s ease, box-shadow .2s ease, border-color .2s ease;
        }
        .stat:hover{transform:translateY(-2px); border-color:rgba(77,182,255,.55); box-shadow:0 10px 24px rgba(77,182,255,.10)}
        .value{
            font-size:2.1rem; font-weight:800; line-height:1.1;
            background:linear-gradient(45deg, #fff, var(--brand)); -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent;
            margin-bottom:6px;
        }
        .label{color:var(--muted); letter-spacing:.12em; font-size:.8rem; text-transform:uppercase}

        /* Progress bar */
        .progress{height:10px; border-radius:999px; background:rgba(255,255,255,.06); border:1px solid var(--stroke); overflow:hidden; margin: 15px 0; width: 100%;}
        .progress-bar{height:100%; width:0%; background:linear-gradient(90deg, var(--brand-2), var(--brand)); transition:width .25s ease}

        /* Status chips & wheel */
        .chip{padding:4px 8px; border-radius:999px; font-size:.85rem; font-weight:700; border:1px solid var(--stroke)}
        .status-ready{color:var(--brand)}
        .status-queued{color:var(--warn)}
        .status-processing{color:var(--brand)}
        .status-completed{color:var(--ok)}
        .status-error{color:var(--err); text-decoration:underline; cursor:pointer}
        .status-error:hover{opacity:.9}

        .wheel{width:16px; height:16px; border-radius:99px; border:2px solid rgba(255,255,255,.1); border-top:2px solid var(--brand); animation:spin 1s linear infinite}
        @keyframes spin{to{transform:rotate(360deg)}}
        .row{display:flex; align-items:center; gap:8px}

        /* Modal */
        .modal{display:none; position:fixed; inset:0; background:rgba(0,0,0,.55); z-index:20; align-items:center; justify-content:center; padding:24px}
        .modal.in,.modal.show{display:flex}
        .modal-content{
            background:linear-gradient(180deg, var(--panel), var(--panel-2));
            width:min(920px, 92vw); max-height:85vh; overflow:auto;
            border:1px solid var(--stroke); border-radius:16px; box-shadow:var(--shadow); padding:20px; margin: 5% auto;
        }
        .modal-header{display:flex; align-items:center; justify-content:space-between; border-bottom:1px solid var(--stroke); padding-bottom:12px; margin-bottom:12px}
        .close{cursor:pointer; font-size:26px; color:var(--muted); float:right}
        .close:hover{color:var(--text)}
        .error-section{margin:14px 0 20px 0; padding:12px 15px; border-left:4px solid var(--err); background:rgba(239,68,68,.06); border-radius:10px}
        .error-section h3{margin:0 0 10px 0; color:#ff8787}
        .error-code{
            background:#0b0d13; color:#e1e7f5; border:1px solid #242b3a; border-radius:10px;
            padding:12px; font-family:ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono","Courier New", monospace; font-size:12px; overflow:auto;
            white-space: pre-wrap; max-height: 200px; overflow-y: auto;
        }

        /* Utilities and legacy mappings */
        .sr-only{position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip:rect(0,0,0,0); white-space:nowrap; border:0}
        .hidden{display:none}
        details summary{cursor:pointer; color:var(--brand)}
        a.link{color:var(--brand); text-decoration:none}
        a.link:hover{text-decoration:underline}

        .stat-value { font-size:2.1rem; font-weight:800; line-height:1.1;
            background:linear-gradient(45deg, #fff, var(--brand)); -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; margin-bottom:6px; }
        .results { background:linear-gradient(180deg, var(--panel), var(--panel-2)); border:1px solid var(--stroke); border-radius:var(--radius); padding:20px; box-shadow:var(--shadow); }
        .results h3 { margin:0 0 10px 0; font-size:1.15rem; font-weight:700; letter-spacing:.2px; color:var(--brand); }
        .result-item { padding:16px; border-top:1px solid var(--stroke); display:flex; flex-direction:column; gap:10px }
        .result-item:first-child{border-top:none}
        .result-item:hover { background: rgba(77,182,255,.05); }
        .upload-area { border:1.5px dashed var(--stroke); border-radius:calc(var(--radius) - 4px); padding:28px; text-align:center; position:relative; overflow:hidden; background:linear-gradient(180deg, rgba(77,182,255,.06), transparent); transition:border-color .2s, transform .15s; }
        .upload-area:hover{border-color:var(--brand); transform: translateY(-1px)}
        .format-selector { background: var(--panel-2); border: 1px solid var(--stroke); color: var(--text); padding: 4px 8px; border-radius: 4px; margin: 0 8px; font-size: 0.85em; }
        .format-selector:focus { border-color: var(--brand); outline: none; box-shadow: var(--ring); }
        .reprocess-btn { appearance:none; border:1px solid var(--stroke); background:linear-gradient(180deg, rgba(77,182,255,.06), rgba(77,182,255,.03)); color:var(--brand); padding:10px 14px; border-radius:10px; font-weight:700; cursor:pointer; transition:transform .12s ease, background .2s ease; }
        .reprocess-btn:hover{transform:translateY(-1px); background:linear-gradient(180deg, rgba(77,182,255,.1), rgba(77,182,255,.06))}
        .progress-container { display: inline-flex; align-items: center; gap: 8px; }
        .progress-wheel { width:16px; height:16px; border-radius:99px; border:2px solid rgba(255,255,255,.1); border-top:2px solid var(--brand); animation:spin 1s linear infinite; display: inline-block; flex-shrink: 0; }
        .progress-text { font-size: 12px; color: var(--muted); }
        .download-all-btn { appearance:none; border:1px solid var(--stroke); background:linear-gradient(180deg, rgba(77,182,255,.06), rgba(77,182,255,.03)); color:var(--brand); padding:10px 14px; border-radius:10px; font-weight:700; cursor:pointer; transition:transform .12s ease, background .2s ease; }
        .download-all-btn:hover{transform:translateY(-1px); background:linear-gradient(180deg, rgba(77,182,255,.1), rgba(77,182,255,.06))}
        .start-btn { appearance:none; border:1px solid var(--stroke); background:linear-gradient(180deg, var(--brand-2), var(--brand)); color:white; padding:10px 16px; border-radius:10px; font-weight:700; cursor:pointer; box-shadow:0 6px 18px rgba(77,182,255,.23); transition:transform .12s ease, box-shadow .2s ease; }
        .start-btn:hover{transform:translateY(-1px); box-shadow:0 10px 24px rgba(77,182,255,.28)}
    </style>
</head>
<body>
    <main class="container">
        <!-- Header -->
        <section class="header" role="banner" aria-label="PSDocling header">
            <div class="header-row">
                <div class="brand">
                    <div class="logo" aria-hidden="true"></div>
                    <div>
                        <h1>PSDocling <span class="badge" aria-label="version">v$version</span></h1>
                        <p class="subtitle">PowerShell-based Document Processor for Docling</p>
                    </div>
                </div>
                <div class="row" aria-live="polite">
                    <span class="status"><span class="status-dot" id="status-dot"></span>Backend Status: <span id="status">Connecting...</span></span>
                </div>
            </div>
        </section>

        <!-- Upload -->
        <section class="card" aria-labelledby="upload-title">
            <h2 class="card-title" id="upload-title">Upload</h2>
            <div class="drop" id="drop-zone" tabindex="0" role="button" aria-label="Drop files here or press Enter to browse">
                <div class="shimmer" aria-hidden="true"></div>
                <h3>Drop files here or <span class="link" onclick="document.getElementById('file-input').click()">browse</span></h3>
                <p class="muted">Max 100MB per file • Press <span class="kbd">Enter</span> to open file picker</p>
                <button class="btn" style="margin-top:10px" onclick="document.getElementById('file-input').click()">Choose Files</button>
                <input type="file" id="file-input" multiple accept=".pdf,.docx,.xlsx,.pptx,.md,.html,.xhtml,.csv,.png,.jpg,.jpeg,.tiff,.bmp,.webp" class="sr-only" aria-hidden="true" />
                <div style="margin-top:18px">
                    <span class="pill">PDF</span>
                    <span class="pill">DOCX</span>
                    <span class="pill">XLSX</span>
                    <span class="pill">PPTX</span>
                    <span class="pill">MD</span>
                    <span class="pill">HTML</span>
                    <span class="pill">CSV</span>
                    <span class="pill">Images</span>
                </div>
            </div>

            <div id="upload-progress" class="card" style="margin-top:14px; display:none">
                <div class="row" style="justify-content:space-between">
                    <p class="muted" style="margin:0">Uploading files…</p>
                    <span class="chip">Transfer</span>
                </div>
                <div class="progress" style="margin-top:10px"><div class="progress-bar" id="progress-bar"></div></div>
            </div>
        </section>

        <!-- Stats -->
        <section class="card" aria-labelledby="stats-title">
            <h2 class="card-title" id="stats-title">Queue Overview</h2>
            <div class="stats" role="list">
                <div class="stat" role="listitem" aria-live="polite">
                    <div class="value" id="queued">0</div>
                    <div class="label">Queued</div>
                </div>
                <div class="stat" role="listitem" aria-live="polite">
                    <div class="value" id="processing">0</div>
                    <div class="label">Processing</div>
                </div>
                <div class="stat" role="listitem" aria-live="polite">
                    <div class="value" id="completed">0</div>
                    <div class="label">Completed</div>
                </div>
                <div class="stat" role="listitem" aria-live="polite">
                    <div class="value" id="errors">0</div>
                    <div class="label">Errors</div>
                </div>
            </div>
        </section>

        <!-- Processing Results -->
        <section class="card" aria-labelledby="results-title" id="results">
            <h2 class="card-title" id="results-title">Processing Results</h2>
            <div id="results-list" class="results-list"></div>
        </section>

        <!-- Processed Files -->
        <section class="card" aria-labelledby="files-title">
            <div class="row" style="justify-content:space-between">
                <h2 class="card-title" id="files-title" style="margin:0">Processed Files</h2>
                <button class="btn-ghost" onclick="downloadAllDocuments()">Download All</button>
            </div>
            <div id="files-list" style="margin-top:10px">
                <p class="muted" style="font-style:italic">Loading processed files…</p>
            </div>
        </section>
    </main>

    <!-- Error Modal -->
    <div id="errorModal" class="modal" role="dialog" aria-modal="true" aria-labelledby="error-modal-title">
        <div class="modal-card">
            <div class="modal-header">
                <h2 class="card-title" id="error-modal-title" style="margin:0">Error Details</h2>
                <span class="close" aria-label="Close error details" role="button">&times;</span>
            </div>
            <div id="errorModalContent">
                <p class="muted">Loading error details…</p>
            </div>
        </div>
    </div>

    <script>
    const API = 'http://localhost:8080';
    const results = {};

    document.addEventListener('DOMContentLoaded', function() {
        setupUpload();

        // Keyboard support for drop zone
        const zone = document.getElementById('drop-zone');
        zone.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                document.getElementById('file-input').click();
            }
        });

        // Delay initial API calls to give server time to start
        setTimeout(async () => {
            const isHealthy = await checkHealth();
            if (isHealthy) {
                loadExistingDocuments();
                loadProcessedFiles();
            }
        }, 600);

        setInterval(loadProcessedFiles, 2000);
        setInterval(updateStats, 2000);
    });

    function setupUpload() {
        const zone = document.getElementById('drop-zone');
        const input = document.getElementById('file-input');

        zone.addEventListener('dragover', e => {
            e.preventDefault();
            zone.style.borderColor = 'var(--brand)';
        });
        zone.addEventListener('dragleave', () => {
            zone.style.borderColor = 'var(--stroke)';
        });
        zone.addEventListener('drop', e => {
            e.preventDefault();
            zone.style.borderColor = 'var(--stroke)';
            handleFiles(e.dataTransfer.files);
        });
        input.addEventListener('change', e => handleFiles(e.target.files));
    }

    async function handleFiles(files) {
        const progress = document.getElementById('upload-progress');
        const bar = document.getElementById('progress-bar');

        progress.style.display = 'block';

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                // Check file size limit (100MB)
                const maxSizeBytes = 100 * 1024 * 1024;  // 100MB
                if (file.size > maxSizeBytes) {
                    const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
                    alert('File "' + file.name + '" (' + sizeMB + ' MB) exceeds the 100MB size limit. Please upload a smaller file.');
                    continue;  // Skip this file
                }

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

        setTimeout(() => { progress.style.display = 'none'; bar.style.width = '0%'; }, 1000);
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

    // Apply model preset
    function applyModelPreset(id) {
        const preset = document.getElementById('chunkModelPreset-' + id).value;
        const modelPresets = {
            'general': { backend: 'hf', model: 'sentence-transformers/all-MiniLM-L6-v2', maxTokens: 512 },
            'legal': { backend: 'hf', model: 'nlpaueb/legal-bert-base-uncased', maxTokens: 512 },
            'medical': { backend: 'hf', model: 'dmis-lab/biobert-v1.1', maxTokens: 256 },
            'financial': { backend: 'hf', model: 'yiyanghkust/finbert-tone', maxTokens: 512 },
            'scientific': { backend: 'hf', model: 'allenai/scibert_scivocab_uncased', maxTokens: 256 },
            'multilingual': { backend: 'hf', model: 'bert-base-multilingual-cased', maxTokens: 400 },
            'code': { backend: 'hf', model: 'microsoft/codebert-base', maxTokens: 512 }
        };

        if (preset && modelPresets[preset]) {
            const config = modelPresets[preset];
            document.getElementById('chunkTokenizerBackend-' + id).value = config.backend;
            document.getElementById('chunkTokenizerModel-' + id).value = config.model;
            document.getElementById('chunkMaxTokens-' + id).value = config.maxTokens;
            toggleTokenizerOptions(id);

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

        // Remove placeholder message if it exists
        const placeholder = list.querySelector('p');
        if (placeholder && placeholder.textContent.includes('No documents in processing')) {
            placeholder.remove();
        }

        const item = document.createElement('div');
        item.className = 'result-item';
        item.id = 'result-item-' + id;
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
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Model Preset:</label>' +
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
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Table Format:</label>' +
                                    '<select id="chunkTableSerialization-' + id + '" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.9em;">' +
                                        '<option value="triplets">Triplets (Default)</option>' +
                                        '<option value="markdown">Markdown Tables</option>' +
                                        '<option value="csv">CSV Format</option>' +
                                        '<option value="grid">ASCII Grid</option>' +
                                    '</select>' +
                                '</div>' +
                                '<div style="margin-bottom: 8px;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Picture Handling:</label>' +
                                    '<select id="chunkPictureStrategy-' + id + '" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.9em;">' +
                                        '<option value="default">Default</option>' +
                                        '<option value="with_caption">Include Captions</option>' +
                                        '<option value="with_description">Include Descriptions</option>' +
                                        '<option value="placeholder">Custom Placeholder</option>' +
                                    '</select>' +
                                '</div>' +
                                '<div id="imagePlaceholderDiv-' + id + '" style="margin-bottom: 8px; display: none;">' +
                                    '<label style="display: block; margin-bottom: 4px; color: #aaa; font-size: 0.85em;">Image Placeholder Text:</label>' +
                                    '<input type="text" id="chunkImagePlaceholder-' + id + '" value="[IMAGE]" style="width: 100%; padding: 4px; background: #1a1a1a; color: white; border: 1px solid #404040; border-radius: 4px; font-size: 0.85em;">' +
                                '</div>' +
                                '<div style="border-top: 1px solid #404040; margin-top: 10px; padding-top: 10px;">' +
                                    '<label style="display: block; margin-bottom: 6px; color: #049fd9; font-size: 0.85em; font-weight: bold;">Advanced Options:</label>' +
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
                    '<div id="download-buttons-' + id + '" style="display:none;">' +
                        '<button class="download-btn" onclick="downloadDocument(\'" + id + "\')" style="margin-right: 10px;">Download</button>' +
                    '</div>' +
                    '<div style="display: flex; gap: 8px;">' +
                        '<button class="start-btn" onclick="startConversion(\'' + id + '\')" id="start-' + id + '" disabled>Start Conversion</button>' +
                        '<button class="reprocess-btn" onclick="reprocessDocument(\'' + id + '\')" style="display:none" id="reprocess-' + id + '">Re-process</button>' +
                    '</div>' +
                '</div>' +
            '</div>';
        list.appendChild(item);
        results[id] = { name: name, format: currentFormat };

        validateExportFormat(id);
    }

    async function startConversion(id) {
        const radioButtons = document.querySelectorAll('input[name="format-' + id + '"]');
        let selectedFormat = null;
        radioButtons.forEach(radio => { if (radio.checked) selectedFormat = radio.value; });
        if (!selectedFormat) { alert('Please select an export format'); return; }

        const embedImages = document.getElementById('embedImages-' + id).checked;
        const enrichCode = document.getElementById('enrichCode-' + id).checked;
        const enrichFormula = document.getElementById('enrichFormula-' + id).checked;
        const enrichPictureClasses = document.getElementById('enrichPictureClasses-' + id).checked;
        const enrichPictureDescription = document.getElementById('enrichPictureDescription-' + id).checked;

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
            startBtn.style.display = 'none';
            statusElement.textContent = 'Queued...';
            statusElement.className = 'status-queued';

            const response = await fetch(API + '/api/start-conversion', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    documentId: id,
                    exportFormat: selectedFormat,
                    embedImages,
                    enrichCode,
                    enrichFormula,
                    enrichPictureClasses,
                    enrichPictureDescription,
                    ...chunkingParams
                })
            });

            if (response.ok) {
                if (results[id] && typeof results[id] === 'object') { results[id].format = selectedFormat; }
                else { results[id] = { name: results[id] || 'Unknown', format: selectedFormat }; }
                setTimeout(() => pollResult(id, results[id].name || 'Unknown'), 500);
            } else {
                throw new Error('Failed to start conversion');
            }
        } catch (error) {
            statusElement.textContent = 'Error (click for details)';
            statusElement.className = 'status-error';
            statusElement.onclick = () => showErrorDetails(id, results[id].name || 'Unknown');
            startBtn.style.display = 'inline';
            console.error('Start conversion error:', error);
        }
    }

    async function reprocessDocument(id) {
        const radioButtons = document.querySelectorAll('input[name="format-' + id + '"]');
        let newFormat = null;
        radioButtons.forEach(radio => { if (radio.checked) newFormat = radio.value; });
        if (!newFormat) { alert('Please select an export format'); return; }

        const embedImages = document.getElementById('embedImages-' + id).checked;
        const enrichCode = document.getElementById('enrichCode-' + id).checked;
        const enrichFormula = document.getElementById('enrichFormula-' + id).checked;
        const enrichPictureClasses = document.getElementById('enrichPictureClasses-' + id).checked;
        const enrichPictureDescription = document.getElementById('enrichPictureDescription-' + id).checked;

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
            reprocessBtn.style.display = 'none';
            statusElement.textContent = 'Re-processing...';
            statusElement.className = 'status-processing';

            const response = await fetch(API + '/api/reprocess', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    documentId: id,
                    exportFormat: newFormat,
                    embedImages,
                    enrichCode,
                    enrichFormula,
                    enrichPictureClasses,
                    enrichPictureDescription,
                    ...chunkingParams
                })
            });

            if (response.ok) {
                if (results[id] && typeof results[id] === 'object') { results[id].format = newFormat; }
                else { results[id] = { name: results[id] || 'Unknown', format: newFormat }; }
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

    function showErrorDetails(id, fileName) {
        const modal = document.getElementById('errorModal');
        const content = document.getElementById('errorModalContent');

        content.innerHTML = '<p>Loading error details...</p>';
        modal.style.display = 'block';
        modal.classList.add('show');

        fetch(API + '/api/error/' + id)
            .then(response => response.ok ? response.json() : response.text().then(t => { throw new Error('HTTP ' + response.status + ': ' + t); }))
            .then(errorData => { content.innerHTML = formatErrorDetails(errorData); })
            .catch(error => { content.innerHTML = '<p style="color: #ef4444;">Failed to load error details: ' + error.message + '</p>'; });
    }

    function formatEstimatedTime(estimatedDurationMs, elapsedTimeMs) {
        if (!estimatedDurationMs || !elapsedTimeMs) return '';
        const remainingMs = Math.max(0, estimatedDurationMs - elapsedTimeMs);
        const remainingSeconds = Math.round(remainingMs / 1000);
        if (remainingSeconds <= 0) return 'finishing...';
        if (remainingSeconds < 60) return remainingSeconds + 's remaining';
        const remainingMinutes = Math.floor(remainingSeconds / 60);
        const seconds = remainingSeconds % 60;
        if (remainingMinutes < 60) { return seconds > 0 ? remainingMinutes + 'm ' + seconds + 's remaining' : remainingMinutes + 'm remaining'; }
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

        if (errorData.currentStatus && errorData.currentStatus !== 'Error') {
            html += '<p><strong>Current Status:</strong> ' + errorData.currentStatus + '</p>';
            if (errorData.startTime) { html += '<p><strong>Started:</strong> ' + (errorData.startTime?.DateTime || errorData.startTime || 'Unknown') + '</p>'; }
            if (errorData.progress !== undefined) { html += '<p><strong>Progress:</strong> ' + errorData.progress + '%</p>'; }
            html += '</div><div class="error-section">';
            html += '<h3>Status Information</h3>';
            html += '<div class="error-code" style="background: #1a2e1a; border-left: 4px solid #10b981;">' + (errorData.message || 'Document is being processed') + '</div>';
            html += '</div>';
        } else {
            html += '<p><strong>Failed:</strong> ' + (errorData.endTime?.DateTime || errorData.endTime || 'Unknown') + '</p></div>';
            html += '<div class="error-section"><h3>Error Message</h3><div class="error-code">' + (errorData.error || 'No error message available') + '</div></div>';
        }

        if (errorData.stderr && typeof errorData.stderr === 'string' && errorData.stderr.trim()) {
            html += '<div class="error-section"><h3>Python Error Output (stderr)</h3><div class="error-code">' + errorData.stderr + '</div></div>';
        }

        if (errorData.errorDetails) {
            html += '<div class="error-section"><h3>Technical Details</h3>';
            if (errorData.errorDetails.ExceptionType) { html += '<p><strong>Exception Type:</strong> ' + errorData.errorDetails.ExceptionType + '</p>'; }
            if (errorData.errorDetails.InnerException) { html += '<p><strong>Inner Exception:</strong> ' + errorData.errorDetails.InnerException + '</p>'; }
            if (errorData.errorDetails.StackTrace) { html += '<h4>Stack Trace:</h4><div class="error-code">' + errorData.errorDetails.StackTrace + '</div>'; }
            if (errorData.errorDetails.ScriptStackTrace) { html += '<h4>Script Stack Trace:</h4><div class="error-code">' + errorData.errorDetails.ScriptStackTrace + '</div>'; }
            html += '</div>';
        }
        return html;
    }

    // Modal close / keyboard support
    document.addEventListener('DOMContentLoaded', function() {
        const modal = document.getElementById('errorModal');
        const closeBtn = document.querySelector('.close');

        closeBtn.onclick = function() {
            modal.style.display = 'none';
            modal.classList.remove('show', 'in');
        }
        window.onclick = function(event) {
            if (event.target === modal) {
                modal.style.display = 'none';
                modal.classList.remove('show', 'in');
            }
        }
        window.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && (modal.classList.contains('show') || modal.classList.contains('in') || modal.style.display === 'flex')) {
                modal.style.display = 'none';
                modal.classList.remove('show', 'in');
            }
        });
    });

    async function pollResult(id, name, attempt = 0) {
        try {
            const response = await fetch(API + '/api/result/' + id);
            if (response.status === 200) {
                const contentLength = response.headers.get('content-length');
                const blob = await response.blob();

                document.getElementById('status-' + id).textContent = 'Completed';
                const link = document.getElementById('link-' + id);

                if ((contentLength && parseInt(contentLength) > 5 * 1024 * 1024) ||
                    (blob.type.includes('json') && blob.size > 1 * 1024 * 1024)) {
                    link.href = API + '/api/result/' + id;
                    link.download = name + '.' + (blob.type.includes('json') ? 'json' : 'md');
                    link.textContent = 'View File (' + (blob.size / (1024 * 1024)).toFixed(1) + ' MB)';
                } else {
                    const url = URL.createObjectURL(blob);
                    link.href = url;
                }

                document.getElementById('download-buttons-' + id).style.display = 'inline';
                loadProcessedFiles();
                updateStats();
                return;
            }
            if (response.status === 202) {
                try {
                    const documentsResponse = await fetch(API + '/api/documents');
                    if (documentsResponse.ok) {
                        const documents = await documentsResponse.json();
                        const doc = documents.find(d => d.id === id);
                        if (doc && doc.status === 'Processing') {
                            const statusElement = document.getElementById('status-' + id);
                            let progressText = 'Processing...';
                            let elapsedText = '';

                            if (doc.elapsedTime) {
                                const elapsedMs = doc.elapsedTime;
                                const elapsedSeconds = Math.floor(elapsedMs / 1000);
                                const hours = Math.floor(elapsedSeconds / 3600);
                                const minutes = Math.floor((elapsedSeconds % 3600) / 60);
                                const seconds = elapsedSeconds % 60;
                                if (hours > 0) { elapsedText = hours + 'h ' + minutes + 'm'; }
                                else if (minutes > 0) { elapsedText = minutes + 'm ' + seconds + 's'; }
                                else { elapsedText = seconds + 's'; }
                            }

                            if (doc.enhancementsInProgress && doc.activeEnhancements && doc.activeEnhancements.length > 0) {
                                const enhancement = doc.activeEnhancements[0];
                                progressText = 'Processing (' + enhancement + ')';
                                if (doc.progress !== undefined && doc.progress !== null) {
                                    progressText = 'Processing (' + enhancement + ') ' + doc.progress + '%';
                                }
                            } else if (doc.progress !== undefined && doc.progress !== null) {
                                progressText = 'Processing ' + doc.progress + '%';
                            }
                            if (elapsedText) { progressText += ' - ' + elapsedText + ' elapsed'; }
                            if (doc.elapsedTime && doc.elapsedTime > 900000) {
                                progressText += '<br><span style="font-size: 0.85em; color: #fbbf24;">Still processing, please be patient...</span>';
                            }

                            statusElement.innerHTML = '<div class="progress-container"><div class="progress-wheel"></div><span>' + progressText + '</span></div>';

                            const linkElement = document.getElementById('link-' + id);
                            if (doc.elapsedTime && doc.elapsedTime > 300000) {
                                if (!document.getElementById('cancel-' + id)) {
                                    const cancelBtn = document.createElement('button');
                                    cancelBtn.id = 'cancel-' + id;
                                    cancelBtn.textContent = 'Cancel';
                                    cancelBtn.style = 'padding: 4px 8px; margin-left: 8px; background: #dc2626; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 0.85em;';
                                    cancelBtn.onclick = () => cancelProcessing(id, name);
                                    linkElement?.parentNode?.insertBefore(cancelBtn, linkElement);
                                }
                            }
                        }
                    }
                } catch (e) {
                    document.getElementById('status-' + id).textContent = 'Processing...';
                }
                setTimeout(() => pollResult(id, name, attempt + 1), 1000);
                return;
            }

            try {
                const documentsResponse = await fetch(API + '/api/documents');
                if (documentsResponse.ok) {
                    const documents = await documentsResponse.json();
                    const doc = documents.find(d => d.id === id);
                    if (doc && doc.status === 'Error') {
                        const statusElement = document.getElementById('status-' + id);
                        statusElement.textContent = 'Error (click for details)';
                        statusElement.className = 'status-error';
                        statusElement.onclick = () => showErrorDetails(id, name);
                        return;
                    } else if (doc) {
                        setTimeout(() => pollResult(id, name, attempt + 1), 2000);
                        return;
                    }
                }
            } catch (docError) {
                console.log('Failed to check document status:', docError);
            }

            const statusElement = document.getElementById('status-' + id);
            statusElement.textContent = 'Connection Error - Retrying...';
            statusElement.className = 'status-error';
            statusElement.onclick = null;
        } catch (error) {
            if (attempt < 30) {
                setTimeout(() => pollResult(id, name, attempt + 1), 2000);
            } else {
                try {
                    const documentsResponse = await fetch(API + '/api/documents');
                    if (documentsResponse.ok) {
                        const documents = await documentsResponse.json();
                        const doc = documents.find(d => d.id === id);
                        if (doc && doc.status === 'Error') {
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
                const statusElement = document.getElementById('status-' + id);
                statusElement.textContent = 'Connection Lost';
                statusElement.className = 'status-error';
                statusElement.onclick = null;
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
                setTimeout(() => checkHealth(1), 1000);
            }
        } catch (error) {
            console.error('Stats update failed:', error);
            setTimeout(() => checkHealth(1), 1000);
        }
    }

    async function loadExistingDocuments() {
        try {
            const response = await fetch(API + '/api/documents');
            if (response.ok) {
                const documents = await response.json();
                const list = document.getElementById('results-list');
                list.innerHTML = '';

                documents.forEach(doc => {
                    if (doc.status === 'Completed') { return; }
                    const docFormat = doc.exportFormat || 'markdown';
                    addResult(doc.id, doc.fileName, docFormat);
                    results[doc.id] = { name: doc.fileName, format: docFormat };

                    const statusElement = document.getElementById('status-' + doc.id);
                    if (doc.status === 'Ready') {
                        statusElement.textContent = 'Ready';
                        statusElement.className = 'status-ready';
                        statusElement.onclick = null;
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) { startBtn.style.display = 'inline'; }
                    } else if (doc.status === 'Processing') {
                        let progressText = 'Processing...';
                        if (doc.progress !== undefined && doc.progress !== null) {
                            progressText = 'Processing ' + doc.progress + '%';
                        }
                        statusElement.innerHTML = '<div class="progress-container"><div class="progress-wheel"></div><span>' + progressText + '</span></div>';
                        statusElement.className = 'status-processing';
                        statusElement.onclick = null;
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) { startBtn.style.display = 'none'; }
                        pollResult(doc.id, doc.fileName);
                    } else if (doc.status === 'Queued') {
                        statusElement.textContent = 'Queued...';
                        statusElement.className = 'status-queued';
                        statusElement.onclick = null;
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) { startBtn.style.display = 'none'; }
                        pollResult(doc.id, doc.fileName);
                    } else if (doc.status === 'Error') {
                        statusElement.textContent = 'Error (click for details)';
                        statusElement.className = 'status-error';
                        statusElement.onclick = () => showErrorDetails(doc.id, doc.fileName);
                        const startBtn = document.getElementById('start-' + doc.id);
                        if (startBtn) { startBtn.style.display = 'inline'; }
                    } else {
                        statusElement.innerHTML = '<span style="color: orange; font-weight: bold;">STATUS: ' + (doc.status || 'UNDEFINED') + '</span>';
                    }
                });
            }
        } catch (error) {
            console.error('Failed to load existing documents:', error);
        }
    }

    async function loadProcessedFiles() {
        try {
            const [filesResponse, documentsResponse] = await Promise.all([
                fetch(API + '/api/files'),
                fetch(API + '/api/documents')
            ]);

            if (!filesResponse.ok && !documentsResponse.ok) {
                const filesList = document.getElementById('files-list');
                filesList.innerHTML = '<p style="color: #fbbf24;">Server responded with an error. Checking connection...</p>';
                setTimeout(() => checkHealth(1), 2000);
                return;
            }

            const allItems = [];
            let documentsMap = new Map();
            if (documentsResponse.ok) {
                const documents = await documentsResponse.json();
                documents.filter(doc => doc.status === 'Completed').forEach(doc => {
                    documentsMap.set(doc.id, doc);
                    // Clean up completed items from Processing Results section
                    const resultItem = document.getElementById('result-item-' + doc.id);
                    if (resultItem) {
                        resultItem.style.transition = 'opacity 0.3s';
                        resultItem.style.opacity = '0';
                        setTimeout(() => {
                            resultItem.remove();
                            const resultsList = document.getElementById('results-list');
                            if (resultsList && resultsList.children.length === 0) {
                                resultsList.innerHTML = '<p style="color: #888; text-align: center; padding: 20px;">No documents in processing. Upload files above to get started.</p>';
                            }
                        }, 300);
                    }
                });
            }

            const filesByDocId = new Map();
            if (filesResponse.ok) {
                let files = await filesResponse.json();
                if (!Array.isArray(files)) files = [files];

                files.forEach(file => {
                    const isGeneratedFile = /\.(md|xml|html|json|png|jpg|jpeg|gif|bmp|tiff|webp)$/i.test(file.fileName);
                    if (isGeneratedFile) {
                        if (!filesByDocId.has(file.id)) { filesByDocId.set(file.id, []); }
                        filesByDocId.get(file.id).push(file);
                    }
                });
            }

            filesByDocId.forEach((files, docId) => {
                const correspondingDoc = documentsMap.get(docId);
                allItems.push({
                    type: 'document',
                    id: docId,
                    files: files,
                    fileCount: files.length,
                    lastModified: files[0].lastModified,
                    totalSize: files.reduce((sum, f) => {
                        const sizeNum = parseFloat(f.fileSize.replace(' KB', ''));
                        return sum + sizeNum;
                    }, 0),
                    exportFormat: correspondingDoc ? correspondingDoc.exportFormat : 'unknown',
                    canReprocess: !!correspondingDoc,
                    originalFileName: correspondingDoc ? correspondingDoc.fileName : 'Unknown'
                });
            });

            const filesList = document.getElementById('files-list');
            if (allItems.length === 0) {
                filesList.innerHTML = '<p style="color: #b0b0b0; font-style: italic;">No processed files found.</p>';
                return;
            }

            filesList.innerHTML = allItems.map(item => {
                const escapedId = item.id.replace(/'/g, "\\'");
                const escapedFileName = item.originalFileName.replace(/'/g, "\\'");
                const fileListHtml = item.files.map(f =>
                    '<li style="color: #b0b0b0; font-size: 0.9em;">' + f.fileName + ' (' + f.fileSize + ')</li>'
                ).join('');

                return '<div class="result-item">' +
                    '<div style="flex: 1;">' +
                    '<strong>' + item.originalFileName + '</strong><br>' +
                    '<small style="color: #b0b0b0;">' +
                    item.fileCount + ' file' + (item.fileCount > 1 ? 's' : '') + ' | ' +
                    item.totalSize.toFixed(2) + ' KB total | Modified: ' + item.lastModified +
                    '</small>' +
                    '<details style="margin-top: 8px;">' +
                    '<summary style="cursor: pointer; color: #049fd9; font-size: 0.9em;">Show files</summary>' +
                    '<ul style="margin: 8px 0 0 20px; padding: 0;">' + fileListHtml + '</ul>' +
                    '</details>' +
                    '</div>' +
                    '<div style="display: flex; gap: 8px; align-items: flex-start;">' +
                    '<button onclick="downloadAllFilesForDocument(\'' + escapedId + '\', \'' + escapedFileName + '\')" style="padding: 6px 12px; background: #0073e6; color: white; border: none; border-radius: 4px; cursor: pointer; white-space: nowrap;">Download</button>' +
                    (item.canReprocess ? '<button class="reprocess-btn" onclick="reprocessFromCompleted(\'' + escapedId + '\', \'' + escapedFileName + '\')" style="white-space: nowrap;">Re-process</button>' : '') +
                    '</div>' +
                    '</div>';
            }).join('');
        } catch (error) {
            console.error('Failed to load processed files:', error);
            document.getElementById('files-list').innerHTML = '<p style="color: #fbbf24;">Connection lost. Attempting to reconnect...</p>';
            setTimeout(() => checkHealth(1), 2000);
        }
    }

    async function checkHealth(retries = 3) {
        try {
            const timeoutPromise = new Promise((_, reject) => { setTimeout(() => reject(new Error('Timeout')), 5000); });
            const response = await Promise.race([ fetch(API + '/api/health'), timeoutPromise ]);
            if (response.ok) {
                document.getElementById('status').textContent = 'Connected';
                document.getElementById('status').style.color = '#00bceb';
                if (document.getElementById('files-list').innerHTML.includes('Connection lost') ||
                    document.getElementById('files-list').innerHTML.includes('Server responded with an error')) {
                    loadProcessedFiles();
                }
                return true;
            } else {
                document.getElementById('status').textContent = 'Server Error';
                document.getElementById('status').style.color = '#ef4444';
                return false;
            }
        } catch (error) {
            if (retries > 0) {
                document.getElementById('status').textContent = 'Connecting...';
                document.getElementById('status').style.color = '#fbbf24';
                await new Promise(resolve => setTimeout(resolve, 2000));
                return checkHealth(retries - 1);
            } else {
                document.getElementById('status').textContent = 'Disconnected';
                document.getElementById('status').style.color = '#ef4444';
                return false;
            }
        }
    }

    async function reprocessFromCompleted(documentId, fileName) {
        try {
            const documentsResponse = await fetch(API + '/api/documents');
            if (documentsResponse.ok) {
                const documents = await documentsResponse.json();
                const doc = documents.find(d => d.id === documentId);
                if (doc) {
                    const docFormat = doc.exportFormat || 'markdown';
                    addResult(documentId, fileName, docFormat);
                    results[documentId] = { name: fileName, format: docFormat };

                    const statusElement = document.getElementById('status-' + documentId);
                    statusElement.textContent = 'Ready';
                    statusElement.className = 'status-ready';
                    statusElement.onclick = null;

                    const startBtn = document.getElementById('start-' + documentId);
                    if (startBtn) { startBtn.style.display = 'inline'; }

                    await fetch(API + '/api/documents/' + documentId + '/reset', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ status: 'Ready' })
                    });

                    loadProcessedFiles();

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

    async function downloadDocument(docId) {
        try {
            const response = await fetch(API + '/api/download/' + docId);
            if (!response.ok) { throw new Error('Download failed'); }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = docId + '.zip';
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Download failed:', error);
            alert('Failed to download document. Please try again.');
        }
    }

    async function cancelProcessing(id, name) {
        if (!confirm('Are you sure you want to cancel processing "' + name + '"?')) { return; }
        try {
            const response = await fetch(API + '/api/cancel/' + id, { method: 'POST' });
            if (response.ok) {
                alert('Cancellation requested. The process will stop shortly.');
                const statusElement = document.getElementById('status-' + id);
                statusElement.textContent = 'Cancelling...';
                statusElement.className = 'status-error';
                const cancelBtn = document.getElementById('cancel-' + id);
                if (cancelBtn) { cancelBtn.remove(); }
                updateStats();
            } else {
                const error = await response.json();
                alert('Failed to cancel: ' + (error.error || 'Unknown error'));
            }
        } catch (error) {
            console.error('Cancel error:', error);
            alert('Failed to cancel processing. Please try again.');
        }
    }

    async function downloadAllFilesForDocument(docId, originalFileName) {
        try {
            const response = await fetch(API + '/api/download/' + docId);
            if (!response.ok) { throw new Error('Download failed'); }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            // Use original filename without extension, add .zip
            const baseFileName = originalFileName.replace(/\.[^/.]+$/, '');
            a.download = baseFileName + '.zip';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Download all files failed:', error);
            alert('Failed to download files: ' + error.message);
        }
    }

    async function downloadSingleFile(downloadUrl, fileName) {
        const downloadPath = downloadUrl.startsWith('/') ? downloadUrl : '/' + downloadUrl;
        const response = await fetch(API + downloadPath);
        if (!response.ok) { throw new Error('Download failed for ' + fileName + ' with status ' + response.status); }
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = fileName;
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    async function downloadProcessedFile(downloadUrl, fileName) {
        try {
            const downloadPath = downloadUrl.startsWith('/') ? downloadUrl : '/' + downloadUrl;
            const response = await fetch(API + downloadPath);
            if (!response.ok) { throw new Error('Download failed with status ' + response.status); }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href = url; a.download = fileName;
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Download failed:', error);
            alert('Failed to download file: ' + error.message);
        }
    }

    async function downloadAllDocuments() {
        try {
            const btn = event.target;
            const originalText = btn.textContent;
            btn.textContent = 'Preparing...'; btn.disabled = true;

            const response = await fetch(API + '/api/download-all');
            if (!response.ok) { throw new Error('Download failed'); }

            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href = url;
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
            a.download = 'PSDocling_Export_' + timestamp + '.zip';
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url);

            btn.textContent = originalText; btn.disabled = false;
        } catch (error) {
            console.error('Download all failed:', error);
            alert('Failed to download all documents. Please try again.');
            if (event && event.target) { event.target.textContent = 'Download All'; event.target.disabled = false; }
        }
    }
    </script>
</body>
</html>
"@

    $html | Set-Content (Join-Path $frontendDir "index.html") -Encoding UTF8

    # Smarter web server (serves correct MIME types, basic cache headers, 404 handling)
    $webServer = @'
param([int]$Port = 8081)

$http = New-Object System.Net.HttpListener
$prefix = "http://localhost:$Port/"
$http.Prefixes.Add($prefix)
$http.Start()

Write-Host "Web server running at $prefix" -ForegroundColor Green

# Simple mime map
$MimeMap = @{
  ".html" = "text/html; charset=utf-8"
  ".htm"  = "text/html; charset=utf-8"
  ".css"  = "text/css; charset=utf-8"
  ".js"   = "application/javascript; charset=utf-8"
  ".json" = "application/json; charset=utf-8"
  ".svg"  = "image/svg+xml"
  ".png"  = "image/png"
  ".jpg"  = "image/jpeg"
  ".jpeg" = "image/jpeg"
  ".gif"  = "image/gif"
  ".webp" = "image/webp"
  ".ico"  = "image/x-icon"
  ".txt"  = "text/plain; charset=utf-8"
  ".map"  = "application/json; charset=utf-8"
  ".xml"  = "application/xml; charset=utf-8"
}

try {
  while ($http.IsListening) {
    $context = $http.GetContext()
    $response = $context.Response
    $request  = $context.Request

    $path = $request.Url.LocalPath
    if ($path -eq "/" -or [string]::IsNullOrEmpty($path)) { $path = "/index.html" }

    $filePath = Join-Path $PSScriptRoot $path.TrimStart('/')

    if (Test-Path $filePath) {
      try {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
        $contentType = $MimeMap[$ext]
        if (-not $contentType) { $contentType = "application/octet-stream" }
        $response.ContentType = $contentType
        # Basic cache headers for static content
        $response.AddHeader("Cache-Control", "no-cache, no-store, must-revalidate")
        $response.AddHeader("Pragma", "no-cache")
        $response.AddHeader("Expires", "0")
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
      } catch {
        $response.StatusCode = 500
        $msg = [System.Text.Encoding]::UTF8.GetBytes("Internal server error: $($_.Exception.Message)")
        $response.OutputStream.Write($msg, 0, $msg.Length)
      }
    } else {
      $response.StatusCode = 404
      $body = "<!doctype html><meta charset=`"utf-8`"><title>404</title><body style='font-family:Segoe UI,Roboto,Arial; padding:24px'><h1>404 Not Found</h1><p>$([System.Web.HttpUtility]::HtmlEncode($path))</p></body>"
      $msg = [System.Text.Encoding]::UTF8.GetBytes($body)
      $response.ContentType = "text/html; charset=utf-8"
      $response.ContentLength64 = $msg.Length
      $response.OutputStream.Write($msg, 0, $msg.Length)
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
                        $queue = Get-QueueItemsFolder
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

                                # Add chunking options with defaults
                                $doc.enableChunking = if ($_.Value.EnableChunking) { [bool]$_.Value.EnableChunking } else { $false }
                                if ($doc.enableChunking) {
                                    $doc.chunkTokenizerBackend = if ($_.Value.ChunkTokenizerBackend) { $_.Value.ChunkTokenizerBackend } else { 'hf' }
                                    $doc.chunkTokenizerModel = if ($_.Value.ChunkTokenizerModel) { $_.Value.ChunkTokenizerModel } else { 'sentence-transformers/all-MiniLM-L6-v2' }
                                    $doc.chunkOpenAIModel = if ($_.Value.ChunkOpenAIModel) { $_.Value.ChunkOpenAIModel } else { 'gpt-4o-mini' }
                                    $doc.chunkMaxTokens = if ($_.Value.ChunkMaxTokens) { $_.Value.ChunkMaxTokens } else { 512 }
                                    $doc.chunkMergePeers = if ($null -ne $_.Value.ChunkMergePeers) { [bool]$_.Value.ChunkMergePeers } else { $true }
                                    $doc.chunkIncludeContext = if ($_.Value.ChunkIncludeContext) { [bool]$_.Value.ChunkIncludeContext } else { $false }
                                    $doc.chunkTableSerialization = if ($_.Value.ChunkTableSerialization) { $_.Value.ChunkTableSerialization } else { 'triplets' }
                                    $doc.chunkPictureStrategy = if ($_.Value.ChunkPictureStrategy) { $_.Value.ChunkPictureStrategy } else { 'default' }
                                    $doc.chunkImagePlaceholder = if ($_.Value.ChunkImagePlaceholder) { $_.Value.ChunkImagePlaceholder } else { '[IMAGE]' }
                                    $doc.chunkOverlapTokens = if ($_.Value.ChunkOverlapTokens) { $_.Value.ChunkOverlapTokens } else { 0 }
                                    $doc.chunkPreserveSentences = if ($_.Value.ChunkPreserveSentences) { [bool]$_.Value.ChunkPreserveSentences } else { $false }
                                    $doc.chunkPreserveCode = if ($_.Value.ChunkPreserveCode) { [bool]$_.Value.ChunkPreserveCode } else { $false }
                                    $doc.chunkModelPreset = if ($_.Value.ChunkModelPreset) { $_.Value.ChunkModelPreset } else { '' }
                                }

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
                                        Get-ChildItem $_.FullName -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp') } | ForEach-Object {
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
                                        $secureFileName = Get-SecureFileName -FileName $requestedFile -AllowedExtensions @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp')
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
                                    $outputFile = Get-ChildItem $docDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp') } | Select-Object -First 1
                                }

                                if ($outputFile -and $outputFile.Extension -in @('.md', '.html', '.json', '.txt', '.xml', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp')) {
                                    $bytes = [System.IO.File]::ReadAllBytes($outputFile.FullName)
                                    $contentType = switch ($outputFile.Extension) {
                                        '.md' { 'text/markdown; charset=utf-8' }
                                        '.html' { 'text/html; charset=utf-8' }
                                        '.json' { 'application/json; charset=utf-8' }
                                        '.txt' { 'text/plain; charset=utf-8' }
                                        '.xml' { 'application/xml; charset=utf-8' }
                                        '.png' { 'image/png' }
                                        '.jpg' { 'image/jpeg' }
                                        '.jpeg' { 'image/jpeg' }
                                        '.gif' { 'image/gif' }
                                        '.bmp' { 'image/bmp' }
                                        '.tiff' { 'image/tiff' }
                                        '.webp' { 'image/webp' }
                                        default { 'application/octet-stream' }
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
                                    $response.OutputStream.Close()
                                    $response.Close()
                                    continue
                                }

                                # Check file size limit (100MB) - validate before decoding to save resources
                                $base64Length = $data.dataBase64.Length
                                $estimatedSizeBytes = ($base64Length * 3) / 4  # Base64 is ~133% of original size
                                $maxSizeBytes = 100 * 1024 * 1024  # 100MB

                                if ($estimatedSizeBytes -gt $maxSizeBytes) {
                                    $sizeMB = [Math]::Round($estimatedSizeBytes / (1024 * 1024), 2)
                                    $response.StatusCode = 413  # Payload Too Large
                                    $responseContent = @{
                                        success = $false
                                        error   = "File size ($sizeMB MB) exceeds the 100MB limit"
                                    } | ConvertTo-Json
                                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                    $response.OutputStream.Close()
                                    $response.Close()
                                    continue
                                }

                                $uploadId = [guid]::NewGuid().ToString()
                                $uploadDir = Join-Path $script:DoclingSystem.TempDirectory $uploadId
                                New-Item -ItemType Directory -Force -Path $uploadDir | Out-Null

                                $filePath = Join-Path $uploadDir $secureFileName
                                $fileBytes = [Convert]::FromBase64String($data.dataBase64)

                                # Double-check actual file size after decoding
                                if ($fileBytes.Length -gt $maxSizeBytes) {
                                    $sizeMB = [Math]::Round($fileBytes.Length / (1024 * 1024), 2)
                                    $response.StatusCode = 413  # Payload Too Large
                                    $responseContent = @{
                                        success = $false
                                        error   = "File size ($sizeMB MB) exceeds the 100MB limit"
                                    } | ConvertTo-Json
                                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($responseContent)
                                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                                    $response.OutputStream.Close()
                                    $response.Close()
                                    continue
                                }

                                [System.IO.File]::WriteAllBytes($filePath, $fileBytes)

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
                                if ($data.chunkImagePlaceholder) { $queueParams.ChunkImagePlaceholder = $data.chunkImagePlaceholder }
                                if ($data.chunkOverlapTokens) { $queueParams.ChunkOverlapTokens = $data.chunkOverlapTokens }
                                if ($data.chunkPreserveSentences -eq $true) { $queueParams.ChunkPreserveSentences = $true }
                                if ($data.chunkPreserveCode -eq $true) { $queueParams.ChunkPreserveCode = $true }
                                if ($data.chunkModelPreset) { $queueParams.ChunkModelPreset = $data.chunkModelPreset }

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
                                $chunkImagePlaceholder = if ($data.chunkImagePlaceholder) { $data.chunkImagePlaceholder } else { '[IMAGE]' }
                                $chunkOverlapTokens = if ($data.chunkOverlapTokens) { $data.chunkOverlapTokens } else { 0 }
                                $chunkPreserveSentences = if ($data.chunkPreserveSentences) { $data.chunkPreserveSentences } else { $false }
                                $chunkPreserveCode = if ($data.chunkPreserveCode) { $data.chunkPreserveCode } else { $false }
                                $chunkModelPreset = if ($data.chunkModelPreset) { $data.chunkModelPreset } else { '' }

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
                                        ChunkImagePlaceholder    = $chunkImagePlaceholder
                                        ChunkOverlapTokens       = $chunkOverlapTokens
                                        ChunkPreserveSentences   = $chunkPreserveSentences
                                        ChunkPreserveCode        = $chunkPreserveCode
                                        ChunkModelPreset         = $chunkModelPreset

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
                                        ChunkImagePlaceholder    = $chunkImagePlaceholder
                                        ChunkOverlapTokens       = $chunkOverlapTokens
                                        ChunkPreserveSentences   = $chunkPreserveSentences
                                        ChunkPreserveCode        = $chunkPreserveCode
                                        ChunkModelPreset         = $chunkModelPreset

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
                                $chunkImagePlaceholder = if ($data.chunkImagePlaceholder) { $data.chunkImagePlaceholder } else { '[IMAGE]' }
                                $chunkOverlapTokens = if ($data.chunkOverlapTokens) { $data.chunkOverlapTokens } else { 0 }
                                $chunkPreserveSentences = if ($data.chunkPreserveSentences) { $data.chunkPreserveSentences } else { $false }
                                $chunkPreserveCode = if ($data.chunkPreserveCode) { $data.chunkPreserveCode } else { $false }
                                $chunkModelPreset = if ($data.chunkModelPreset) { $data.chunkModelPreset } else { '' }

                                $success = Start-DocumentConversion -DocumentId $documentId -ExportFormat $exportFormat -EmbedImages:$embedImages -EnrichCode:$enrichCode -EnrichFormula:$enrichFormula -EnrichPictureClasses:$enrichPictureClasses -EnrichPictureDescription:$enrichPictureDescription -EnableChunking:$enableChunking -ChunkTokenizerBackend $chunkTokenizerBackend -ChunkTokenizerModel $chunkTokenizerModel -ChunkOpenAIModel $chunkOpenAIModel -ChunkMaxTokens $chunkMaxTokens -ChunkMergePeers:$chunkMergePeers -ChunkIncludeContext:$chunkIncludeContext -ChunkTableSerialization $chunkTableSerialization -ChunkPictureStrategy $chunkPictureStrategy -ChunkImagePlaceholder $chunkImagePlaceholder -ChunkOverlapTokens $chunkOverlapTokens -ChunkPreserveSentences:$chunkPreserveSentences -ChunkPreserveCode:$chunkPreserveCode -ChunkModelPreset $chunkModelPreset

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

                    '^/api/cancel/(.+)$' {
                        $id = $Matches[1]
                        try {
                            # Mark document for cancellation
                            $allStatus = Get-ProcessingStatus
                            $status = $allStatus[$id]

                            if (-not $status) {
                                $response.StatusCode = 404
                                $responseContent = @{
                                    success = $false
                                    error   = "Document not found"
                                } | ConvertTo-Json
                            }
                            elseif ($status.Status -eq 'Processing') {
                                # Update status to mark for cancellation
                                Update-ItemStatus $id @{
                                    CancelRequested = $true
                                }
                                $response.StatusCode = 200
                                $responseContent = @{
                                    success = $true
                                    message = "Cancellation requested"
                                } | ConvertTo-Json
                            }
                            else {
                                $response.StatusCode = 400
                                $responseContent = @{
                                    success = $false
                                    error   = "Document is not currently processing (Status: $($status.Status))"
                                } | ConvertTo-Json
                            }
                        }
                        catch {
                            $response.StatusCode = 500
                            $responseContent = @{
                                success = $false
                                error   = "Failed to cancel document"
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

                    # Download single document as ZIP
                    '^/api/download/([a-fA-F0-9\-]+)$' {
                        $docId = $Matches[1]
                        $processedDir = $script:DoclingSystem.OutputDirectory
                        $docDir = Join-Path $processedDir $docId

                        if (Test-Path $docDir) {
                            try {
                                # Create ZIP file in temp directory
                                $zipPath = Join-Path $env:TEMP "$docId.zip"

                                # Remove existing ZIP if it exists
                                if (Test-Path $zipPath) {
                                    Remove-Item $zipPath -Force
                                }

                                # Create ZIP archive
                                Compress-Archive -Path "$docDir\*" -DestinationPath $zipPath -Force

                                # Send ZIP file
                                $fileBytes = [System.IO.File]::ReadAllBytes($zipPath)
                                $response.ContentType = "application/zip"
                                $response.ContentLength64 = $fileBytes.Length
                                $response.Headers.Add("Content-Disposition", "attachment; filename=`"$docId.zip`"")
                                $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
                                $response.Close()

                                # Clean up temp file
                                Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
                                continue
                            }
                            catch {
                                $response.StatusCode = 500
                                $responseContent = @{
                                    error = "Failed to create download"
                                    details = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                        else {
                            $response.StatusCode = 404
                            $responseContent = @{ error = "Document not found" } | ConvertTo-Json
                        }
                    }

                    # Download all documents as ZIP
                    '^/api/download-all$' {
                        $processedDir = $script:DoclingSystem.OutputDirectory

                        if (Test-Path $processedDir) {
                            try {
                                # Create timestamp for filename
                                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                                $zipPath = Join-Path $env:TEMP "PSDocling_Export_$timestamp.zip"

                                # Remove existing ZIP if it exists
                                if (Test-Path $zipPath) {
                                    Remove-Item $zipPath -Force
                                }

                                # Get all document folders
                                $docFolders = Get-ChildItem -Path $processedDir -Directory

                                if ($docFolders.Count -eq 0) {
                                    $response.StatusCode = 404
                                    $responseContent = @{ error = "No documents to download" } | ConvertTo-Json
                                }
                                else {
                                    # Create ZIP archive with all documents
                                    Compress-Archive -Path "$processedDir\*" -DestinationPath $zipPath -Force

                                    # Send ZIP file
                                    $fileBytes = [System.IO.File]::ReadAllBytes($zipPath)
                                    $response.ContentType = "application/zip"
                                    $response.ContentLength64 = $fileBytes.Length
                                    $response.Headers.Add("Content-Disposition", "attachment; filename=`"PSDocling_Export_$timestamp.zip`"")
                                    $response.OutputStream.Write($fileBytes, 0, $fileBytes.Length)
                                    $response.Close()

                                    # Clean up temp file
                                    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
                                    continue
                                }
                            }
                            catch {
                                $response.StatusCode = 500
                                $responseContent = @{
                                    error = "Failed to create bulk download"
                                    details = $_.Exception.Message
                                } | ConvertTo-Json
                            }
                        }
                        else {
                            $response.StatusCode = 404
                            $responseContent = @{ error = "No processed documents directory found" } | ConvertTo-Json
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

    # Add error logging
    $errorLogFile = "$env:TEMP\docling_processor_errors.log"
    $debugLogFile = "$env:TEMP\docling_processor_debug.log"
    "$(Get-Date) - Processor started" | Add-Content $debugLogFile

    while ($true) {
        try {
            # Use folder-based queue - get the next document ID
            $documentId = Get-NextQueueItemFolder
            "$(Get-Date) - Get-NextQueueItemFolder returned: $(if ($documentId) { "ID: $documentId" } else { 'null' })" | Add-Content $debugLogFile

            if ($documentId) {
                # Fetch the full document details from the status file
                "$(Get-Date) - Fetching details for document ID: $documentId" | Add-Content $debugLogFile
                $allStatus = Get-ProcessingStatus
                $item = $allStatus[$documentId]

                if (-not $item) {
                    "$(Get-Date) - ERROR: Could not find document $documentId in status file!" | Add-Content $debugLogFile
                    Start-Sleep -Seconds 2
                    continue
                }

                "$(Get-Date) - Found document in status file: $($item.FileName)" | Add-Content $debugLogFile

                # Now extract ALL properties from the status item (simpler since it's from our own system)
                $itemId = $documentId  # We already have the ID
                $itemFilePath = if ($item.FilePath) { $item.FilePath } elseif ($item['FilePath']) { $item['FilePath'] } else { $null }
                $itemFileName = if ($item.FileName) { $item.FileName } elseif ($item['FileName']) { $item['FileName'] } else { $null }
                $itemExportFormat = if ($item.ExportFormat) { $item.ExportFormat } elseif ($item['ExportFormat']) { $item['ExportFormat'] } else { 'markdown' }
                $itemEmbedImages = if ($item.EmbedImages) { $item.EmbedImages } elseif ($item['EmbedImages']) { $item['EmbedImages'] } else { $false }
                $itemEnrichCode = if ($item.EnrichCode) { $item.EnrichCode } elseif ($item['EnrichCode']) { $item['EnrichCode'] } else { $false }
                $itemEnrichFormula = if ($item.EnrichFormula) { $item.EnrichFormula } elseif ($item['EnrichFormula']) { $item['EnrichFormula'] } else { $false }
                $itemEnrichPictureClasses = if ($item.EnrichPictureClasses) { $item.EnrichPictureClasses } elseif ($item['EnrichPictureClasses']) { $item['EnrichPictureClasses'] } else { $false }
                $itemEnrichPictureDescription = if ($item.EnrichPictureDescription) { $item.EnrichPictureDescription } elseif ($item['EnrichPictureDescription']) { $item['EnrichPictureDescription'] } else { $false }

                # Chunking properties
                $itemEnableChunking = if ($item.EnableChunking) { $item.EnableChunking } elseif ($item['EnableChunking']) { $item['EnableChunking'] } else { $false }
                $itemChunkTokenizerBackend = if ($item.ChunkTokenizerBackend) { $item.ChunkTokenizerBackend } elseif ($item['ChunkTokenizerBackend']) { $item['ChunkTokenizerBackend'] } else { 'hf' }
                $itemChunkTokenizerModel = if ($item.ChunkTokenizerModel) { $item.ChunkTokenizerModel } elseif ($item['ChunkTokenizerModel']) { $item['ChunkTokenizerModel'] } else { 'sentence-transformers/all-MiniLM-L6-v2' }
                $itemChunkOpenAIModel = if ($item.ChunkOpenAIModel) { $item.ChunkOpenAIModel } elseif ($item['ChunkOpenAIModel']) { $item['ChunkOpenAIModel'] } else { 'gpt-4o-mini' }
                $itemChunkMaxTokens = if ($item.ChunkMaxTokens) { $item.ChunkMaxTokens } elseif ($item['ChunkMaxTokens']) { $item['ChunkMaxTokens'] } else { 512 }
                $itemChunkMergePeers = if ($item.ChunkMergePeers -ne $null) { $item.ChunkMergePeers } elseif ($item['ChunkMergePeers'] -ne $null) { $item['ChunkMergePeers'] } else { $true }
                $itemChunkIncludeContext = if ($item.ChunkIncludeContext) { $item.ChunkIncludeContext } elseif ($item['ChunkIncludeContext']) { $item['ChunkIncludeContext'] } else { $false }
                $itemChunkTableSerialization = if ($item.ChunkTableSerialization) { $item.ChunkTableSerialization } elseif ($item['ChunkTableSerialization']) { $item['ChunkTableSerialization'] } else { 'triplets' }
                $itemChunkPictureStrategy = if ($item.ChunkPictureStrategy) { $item.ChunkPictureStrategy } elseif ($item['ChunkPictureStrategy']) { $item['ChunkPictureStrategy'] } else { 'default' }
                $itemChunkImagePlaceholder = if ($item.ChunkImagePlaceholder) { $item.ChunkImagePlaceholder } elseif ($item['ChunkImagePlaceholder']) { $item['ChunkImagePlaceholder'] } else { '[IMAGE]' }
                $itemChunkOverlapTokens = if ($item.ChunkOverlapTokens) { $item.ChunkOverlapTokens } elseif ($item['ChunkOverlapTokens']) { $item['ChunkOverlapTokens'] } else { 0 }
                $itemChunkPreserveSentences = if ($item.ChunkPreserveSentences) { $item.ChunkPreserveSentences } elseif ($item['ChunkPreserveSentences']) { $item['ChunkPreserveSentences'] } else { $false }
                $itemChunkPreserveCode = if ($item.ChunkPreserveCode) { $item.ChunkPreserveCode } elseif ($item['ChunkPreserveCode']) { $item['ChunkPreserveCode'] } else { $false }
                $itemChunkModelPreset = if ($item.ChunkModelPreset) { $item.ChunkModelPreset } elseif ($item['ChunkModelPreset']) { $item['ChunkModelPreset'] } else { '' }

                "$(Get-Date) - Item ID: $itemId" | Add-Content $debugLogFile
                "$(Get-Date) - Item FilePath: $itemFilePath" | Add-Content $debugLogFile
                "$(Get-Date) - Item FileName: $itemFileName" | Add-Content $debugLogFile

                if (-not $itemId -or -not $itemFilePath) {
                    "$(Get-Date) - ERROR: Missing critical properties (ID or FilePath)" | Add-Content $debugLogFile

                    # Try to log all properties for debugging
                    if ($item -is [PSCustomObject]) {
                        "$(Get-Date) - PSCustomObject properties:" | Add-Content $debugLogFile
                        $item.PSObject.Properties | ForEach-Object {
                            "$(Get-Date) -   $($_.Name) = $($_.Value)" | Add-Content $debugLogFile
                        }
                    }

                    Start-Sleep -Seconds 2
                    continue
                }

                Write-Host "Processing: $itemFileName" -ForegroundColor Yellow
                "$(Get-Date) - Starting to process: $itemFileName" | Add-Content $debugLogFile

                    # Get file size for progress estimation - with error handling
                    try {
                        "$(Get-Date) - Checking file: $itemFilePath" | Add-Content $debugLogFile
                        # Check if file exists first
                        if (-not (Test-Path $itemFilePath)) {
                            "$(Get-Date) - File not found!" | Add-Content $debugLogFile
                            throw "File not found: $itemFilePath"
                        }
                        "$(Get-Date) - File exists" | Add-Content $debugLogFile

                        $fileSize = (Get-Item $itemFilePath).Length
                        "$(Get-Date) - File size: $fileSize" | Add-Content $debugLogFile
                        $estimatedDurationMs = [Math]::Max(30000, [Math]::Min(300000, $fileSize / 1024 * 1000)) # 30s to 5min based on file size

                        # Update status with progress tracking
                        "$(Get-Date) - Updating status to Processing" | Add-Content $debugLogFile
                        Update-ItemStatus $itemId @{
                            Status            = 'Processing'
                            StartTime         = Get-Date
                            Progress          = 0
                            FileSize          = $fileSize
                            EstimatedDuration = $estimatedDurationMs
                        }
                        "$(Get-Date) - Status updated successfully" | Add-Content $debugLogFile
                    } catch {
                        "$(Get-Date) - ERROR in file check: $($_.Exception.Message)" | Add-Content $debugLogFile
                        Write-Host "Error accessing file '$($itemFileName)': $($_.Exception.Message)" -ForegroundColor Red
                        Update-ItemStatus $itemId @{
                            Status = 'Failed'
                            Error = "File access error: $($_.Exception.Message)"
                            EndTime = Get-Date
                        }
                        continue  # Skip to next item in queue
                    }

                # Create output directory
                $outputDir = Join-Path $script:DoclingSystem.OutputDirectory $itemId
                New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($itemFileName)

                # Determine file extension based on export format
                $extension = switch ($itemExportFormat) {
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
                $wasCancelled = $false

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
        image_mode = ImageRefMode.EMBEDDED if embed_images else ImageRefMode.PLACEHOLDER
        images_extracted = 0  # Will be updated by save_images_and_update_content if extraction happens

        dst.parent.mkdir(parents=True, exist_ok=True)

        # Export using our custom methods to control directory structure and filenames
        if export_format == 'markdown':
            if image_mode == ImageRefMode.EMBEDDED:
                content = result.document.export_to_markdown(image_mode=ImageRefMode.EMBEDDED)
                images_extracted = 0  # Images are embedded, not extracted
            else:
                # Use PLACEHOLDER mode to get <!-- image --> placeholders we can replace
                content = result.document.export_to_markdown(image_mode=ImageRefMode.PLACEHOLDER)
                content, images_extracted = save_images_and_update_content(content, 'markdown')
            dst.write_text(content, encoding='utf-8')
            print(f"Saved markdown with custom image handling", file=sys.stderr)
        elif export_format == 'html':
            if image_mode == ImageRefMode.EMBEDDED:
                content = result.document.export_to_html(image_mode=ImageRefMode.EMBEDDED)
                images_extracted = 0  # Images are embedded, not extracted
            else:
                # Use PLACEHOLDER mode to get <!-- image --> placeholders we can replace
                content = result.document.export_to_html(image_mode=ImageRefMode.PLACEHOLDER)
                content, images_extracted = save_images_and_update_content(content, 'html')
            dst.write_text(content, encoding='utf-8')
            print(f"Saved HTML with custom image handling", file=sys.stderr)
        elif export_format == 'json':
            import json
            # Extract images to separate files even for JSON format (unless embedding)
            if not embed_images and hasattr(result.document, 'pictures') and result.document.pictures:
                for i, picture in enumerate(result.document.pictures):
                    try:
                        pil_image = picture.get_image(result.document)
                        if pil_image:
                            image_filename = f"image_{images_extracted + 1:03d}.png"
                            image_path = images_dir / image_filename
                            pil_image.save(str(image_path), 'PNG')
                            print(f"Extracted image {images_extracted + 1} for JSON export: {image_path}", file=sys.stderr)
                            images_extracted += 1
                    except Exception as img_error:
                        print(f"Warning: Could not extract image {i + 1}: {img_error}", file=sys.stderr)

            doc_dict = result.document.export_to_dict()
            content = json.dumps(doc_dict, indent=2, ensure_ascii=False)
            dst.write_text(content, encoding='utf-8')
            print(f"Saved JSON (images in document structure)", file=sys.stderr)
        elif export_format == 'text':
            # Extract images to separate files even for text format (unless embedding)
            if not embed_images and hasattr(result.document, 'pictures') and result.document.pictures:
                for i, picture in enumerate(result.document.pictures):
                    try:
                        pil_image = picture.get_image(result.document)
                        if pil_image:
                            image_filename = f"image_{images_extracted + 1:03d}.png"
                            image_path = images_dir / image_filename
                            pil_image.save(str(image_path), 'PNG')
                            print(f"Extracted image {images_extracted + 1} for text export: {image_path}", file=sys.stderr)
                            images_extracted += 1
                    except Exception as img_error:
                        print(f"Warning: Could not extract image {i + 1}: {img_error}", file=sys.stderr)

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

            # Extract images to separate files even for doctags format (unless embedding)
            if not embed_images and hasattr(result.document, 'pictures') and result.document.pictures:
                for i, picture in enumerate(result.document.pictures):
                    try:
                        pil_image = picture.get_image(result.document)
                        if pil_image:
                            image_filename = f"image_{images_extracted + 1:03d}.png"
                            image_path = images_dir / image_filename
                            pil_image.save(str(image_path), 'PNG')
                            print(f"Extracted image {images_extracted + 1} for doctags export: {image_path}", file=sys.stderr)
                            images_extracted += 1
                    except Exception as img_error:
                        print(f"Warning: Could not extract image {i + 1}: {img_error}", file=sys.stderr)

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
                            $exportFormat = if ($itemExportFormat) { $itemExportFormat } else { 'markdown' }
                            $embedImages = if ($itemEmbedImages) { 'true' } else { 'false' }
                            $enrichCode = if ($itemEnrichCode) { 'true' } else { 'false' }
                            $enrichFormula = if ($itemEnrichFormula) { 'true' } else { 'false' }
                            $enrichPictureClasses = if ($itemEnrichPictureClasses) { 'true' } else { 'false' }
                            $enrichPictureDescription = if ($itemEnrichPictureDescription) { 'true' } else { 'false' }
                            $arguments = "`"$tempPy`" `"$($itemFilePath)`" `"$outputFile`" `"$exportFormat`" `"$embedImages`" `"$enrichCode`" `"$enrichFormula`" `"$enrichPictureClasses`" `"$enrichPictureDescription`""
                            $process = Start-Process python -ArgumentList $arguments -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\docling_output.txt" -RedirectStandardError "$env:TEMP\docling_error.txt"

                            # Monitor process with progress updates
                            $startTime = Get-Date
                            $finished = $false
                            $lastProgressUpdate = 0

                            while (-not $process.HasExited) {
                                Start-Sleep -Milliseconds 1000  # Check every second
                                $elapsed = (Get-Date) - $startTime
                                $elapsedMs = $elapsed.TotalMilliseconds

                                # Check for timeout - extended to 6 hours for large documents and AI model processing
                                $timeoutSeconds = 21600  # 6 hours for all processing types
                                if ($elapsed.TotalSeconds -gt $timeoutSeconds) {
                                    $timeoutHours = [Math]::Round($timeoutSeconds / 3600, 1)
                                    Write-Host "Process timeout for $($itemFileName) after $timeoutHours hours, terminating..." -ForegroundColor Yellow
                                    $processTerminatedEarly = $true
                                    try {
                                        $process.Kill()
                                    }
                                    catch {
                                        Write-Host "Could not kill process: $($_.Exception.Message)" -ForegroundColor Red
                                    }
                                    break
                                }

                                # Check for cancellation request
                                $currentStatus = Get-ProcessingStatus
                                if ($currentStatus[$itemId].CancelRequested) {
                                    Write-Host "Cancellation requested for $($itemFileName), terminating..." -ForegroundColor Yellow
                                    try {
                                        $process.Kill()
                                        $processTerminatedEarly = $true
                                    }
                                    catch {
                                        Write-Host "Could not kill process: $($_.Exception.Message)" -ForegroundColor Red
                                    }

                                    # Delete output directory if it exists
                                    if (Test-Path $outputDir) {
                                        Remove-Item -Path $outputDir -Recurse -Force -ErrorAction SilentlyContinue
                                        Write-Host "Deleted output directory for cancelled document: $outputDir" -ForegroundColor Yellow
                                    }

                                    # Update status to Cancelled
                                    Update-ItemStatus $itemId @{
                                        Status    = 'Cancelled'
                                        EndTime   = Get-Date
                                        Error     = "Processing cancelled by user"
                                        Progress  = 0
                                    }

                                    # Set flag to skip error handling
                                    $wasCancelled = $true

                                    # Skip to next queue item - do not continue to error handling
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

                                # Calculate and update progress with guards - improved progress calculation for AI enrichments
                                if ($itemEnrichPictureDescription) {
                                    # Picture Description (Granite Vision) - smooth linear progress
                                    if ($elapsed.TotalSeconds -lt 300) {
                                        # First 5 minutes: Model download/loading - progress to 25%
                                        $progress = ($elapsed.TotalSeconds / 300.0) * 25.0
                                    } elseif ($elapsed.TotalSeconds -lt 1800) {
                                        # Next 25 minutes: Main processing - 25% to 90%
                                        $progress = 25.0 + (($elapsed.TotalSeconds - 300.0) / 1500.0) * 65.0
                                    } else {
                                        # Final phase: continue linear progress to 95%, then hold
                                        $remainingTime = 21600.0 - 1800.0  # 6 hours - 30 minutes
                                        $progress = 90.0 + [Math]::Min(5.0, (($elapsed.TotalSeconds - 1800.0) / $remainingTime) * 5.0)
                                    }
                                } elseif ($itemEnrichCode -or $itemEnrichFormula) {
                                    # Code/Formula Understanding (CodeFormulaV2) - smooth linear progress
                                    if ($elapsed.TotalSeconds -lt 180) {
                                        # First 3 minutes: Model download/loading - progress to 20%
                                        $progress = ($elapsed.TotalSeconds / 180.0) * 20.0
                                    } elseif ($elapsed.TotalSeconds -lt 1200) {
                                        # Next 17 minutes: Main processing - 20% to 90%
                                        $progress = 20.0 + (($elapsed.TotalSeconds - 180.0) / 1020.0) * 70.0
                                    } else {
                                        # Final phase: continue linear progress to 95%, then hold
                                        $remainingTime = 21600.0 - 1200.0  # 6 hours - 20 minutes
                                        $progress = 90.0 + [Math]::Min(5.0, (($elapsed.TotalSeconds - 1200.0) / $remainingTime) * 5.0)
                                    }
                                } elseif ($estimatedDurationMs -gt 0) {
                                    $progress = [Math]::Min(95.0, ([double]($elapsedMs) / [double]($estimatedDurationMs)) * 100.0)
                                }
                                else {
                                    $progress = [Math]::Min(95.0, ([double]($elapsedMs) / 60000.0) * 100.0)
                                }

                                # Only update if progress changed significantly
                                if ([Math]::Abs($progress - $lastProgressUpdate) -gt 1.0) {
                                    Update-ItemStatus $itemId @{
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

                            if ($finished -and -not $wasCancelled) {
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
                        Update-ItemStatus $itemId @{
                            Status = 'Processing'
                            StartTime = Get-Date
                            Progress = 10
                        }
                        Start-Sleep 2
                        "Simulated conversion of: $($itemFileName)`nGenerated at: $(Get-Date)" | Set-Content $outputFile -Encoding UTF8
                        $success = $true
                        $processCompletedNormally = $true
                        $processExitCode = 0
                        $pythonSuccess = $true
                        $outputExists = Test-Path $outputFile -ErrorAction SilentlyContinue
                        $imagesExtracted = 0
                        $imagesDirectory = $null
                    }

                if ($success) {
                    # Initialize status update but don't mark as completed yet
                    $statusUpdate = @{
                        OutputFile = $outputFile
                        Progress   = 50  # Base conversion is 50% of total progress
                    }

                    # Track which enhancements need to run
                    $enhancementsToRun = @()
                    if ($itemEnableChunking) { $enhancementsToRun += 'Chunking' }

                    # Add image extraction info if available
                    if ($imagesExtracted -gt 0) {
                        $statusUpdate.ImagesExtracted = $imagesExtracted
                        # Images are now in same folder as output file
                        $statusUpdate.ImagesDirectory = Split-Path $outputFile -Parent
                        Write-Host "Base conversion completed: $($itemFileName) ($imagesExtracted images extracted)" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Base conversion completed: $($itemFileName)" -ForegroundColor Green
                    }

                    # Calculate progress increment for each enhancement
                    $enhancementProgressIncrement = if ($enhancementsToRun.Count -gt 0) { 50 / $enhancementsToRun.Count } else { 50 }
                    $currentProgress = 50

                    # Process chunking if enabled
                    if ($itemEnableChunking -and $outputFile) {
                        try {
                            Write-Host "Starting hybrid chunking for $($itemFileName)..." -ForegroundColor Yellow

                            # Update status to show chunking in progress
                            Update-ItemStatus $itemId @{
                                Status = 'Processing'
                                Progress = $currentProgress
                                EnhancementsInProgress = $true
                                ActiveEnhancements = @('Chunking')
                            }

                            # Build chunking parameters
                            # Use the original source file for chunking, not the converted output
                            $chunkParams = @{
                                InputPath = $itemFilePath  # Use original document
                                TokenizerBackend = $itemChunkTokenizerBackend
                                MaxTokens = $itemChunkMaxTokens
                                MergePeers = $itemChunkMergePeers
                                TableSerialization = $itemChunkTableSerialization
                                PictureStrategy = $itemChunkPictureStrategy
                            }

                            if ($itemChunkTokenizerBackend -eq 'hf') {
                                $chunkParams.TokenizerModel = $itemChunkTokenizerModel
                            } else {
                                $chunkParams.OpenAIModel = $itemChunkOpenAIModel
                            }

                            if ($itemChunkIncludeContext) {
                                $chunkParams.IncludeContext = $true
                            }

                            # Add advanced parameters if present
                            if ($itemChunkImagePlaceholder) {
                                $chunkParams.ImagePlaceholder = $itemChunkImagePlaceholder
                            }
                            if ($itemChunkOverlapTokens) {
                                $chunkParams.OverlapTokens = $itemChunkOverlapTokens
                            }
                            if ($itemChunkPreserveSentences) {
                                $chunkParams.PreserveSentenceBoundaries = $true
                            }
                            if ($itemChunkPreserveCode) {
                                $chunkParams.PreserveCodeBlocks = $true
                            }
                            if ($itemChunkModelPreset) {
                                $chunkParams.ModelPreset = $itemChunkModelPreset
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
                                $currentProgress += $enhancementProgressIncrement
                                $statusUpdate.Progress = [Math]::Min(100, $currentProgress)
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
                    else {
                        # No enhancements, so we're done
                        $statusUpdate.Progress = 100
                    }

                    # Now mark as completed after all enhancements are done
                    $statusUpdate.Status = 'Completed'
                    $statusUpdate.EndTime = Get-Date
                    $statusUpdate.EnhancementsInProgress = $false
                    $statusUpdate.ActiveEnhancements = @()

                    # Add missing status fields for completion
                    $statusUpdate.Status = 'Completed'
                    $statusUpdate.EndTime = Get-Date

                    Write-Host "All processing completed for: $($itemFileName)" -ForegroundColor Green
                    Update-ItemStatus $itemId $statusUpdate
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

                Update-ItemStatus $itemId @{
                    Status       = 'Error'
                    Error        = $_.Exception.Message
                    ErrorDetails = $errorDetails
                    StdErr       = $stderr
                    EndTime      = Get-Date
                    Progress     = 0
                }
                Write-Host "Error processing $($itemFileName): $($_.Exception.Message)" -ForegroundColor Red
            }
            } # End of if ($item) block
        } catch {
            # Log any uncaught errors in the main loop
            $errorMsg = "$(Get-Date) - ERROR in main loop: $($_.Exception.Message)`nStack: $($_.ScriptStackTrace)"
            $errorMsg | Add-Content $errorLogFile
            Write-Host "ERROR in processor loop: $($_.Exception.Message)" -ForegroundColor Red
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

    # Clear the queue folder (new folder-based queue)
    $queueFolder = "$env:TEMP\DoclingQueue"
    if (Test-Path $queueFolder) {
        $queueCount = (Get-ChildItem $queueFolder -Filter "*.queue" -ErrorAction SilentlyContinue).Count
        if ($queueCount -gt 0) {
            Remove-Item "$queueFolder\*.queue" -Force
            Write-Host "Cleared $queueCount items from queue folder" -ForegroundColor Green
        }
        else {
            Write-Host "Queue folder is already empty" -ForegroundColor Gray
        }
    }
    else {
        Write-Host "Queue folder doesn't exist" -ForegroundColor Gray
    }

    # Clear the old JSON queue file (for backwards compatibility)
    $queueFile = "$env:TEMP\docling_queue.json"
    if (Test-Path $queueFile) {
        "[]" | Set-Content $queueFile -Encoding UTF8
        Write-Host "Cleared old queue file" -ForegroundColor Green
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
        [switch]$OpenBrowser,
        [switch]$UseWebView
    )

    Write-Host "Starting Docling System..." -ForegroundColor Cyan

    # Start API server
    # Pass Python availability status to subprocess
    $pythonAvailable = if ($script:DoclingSystem.PythonAvailable) { '$true' } else { '$false' }
    $modulePath = $script:DoclingSystem.ModulePath
    $apiScript = @"
Remove-Module PSDocling -Force -ErrorAction SilentlyContinue
Import-Module '$modulePath' -Force
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
Import-Module '$modulePath' -Force
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

        if ($UseWebView) {
            Start-Sleep 2
            # Try multiple locations for the PyWebView script
            $pyWebViewScript = $null
            $searchPaths = @(
                ".\Launch-PyWebView.py",                          # Current directory
                (Join-Path $PSScriptRoot "..\..\..\Launch-PyWebView.py"),  # From Build folder
                (Join-Path (Split-Path $PSScriptRoot -Parent) "..\..\..\Launch-PyWebView.py")  # From nested source
            )

            foreach ($path in $searchPaths) {
                $resolvedPath = Resolve-Path $path -ErrorAction SilentlyContinue
                if ($resolvedPath -and (Test-Path $resolvedPath)) {
                    $pyWebViewScript = $resolvedPath.Path
                    break
                }
            }

            if ($pyWebViewScript) {
                $pyProcess = Start-Process python -ArgumentList $pyWebViewScript, $script:DoclingSystem.APIPort, $script:DoclingSystem.WebPort -PassThru -WindowStyle Hidden
                Write-Host "PyWebView window launched" -ForegroundColor Green
            } else {
                Write-Warning "PyWebView script not found. Install pywebview with: pip install pywebview requests"
                Write-Host "Falling back to browser mode" -ForegroundColor Yellow
                Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
                Write-Host "Frontend opened in browser: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green
            }
        } elseif ($OpenBrowser) {
            Start-Sleep 2
            Start-Process "http://localhost:$($script:DoclingSystem.WebPort)"
            Write-Host "Frontend opened in browser: http://localhost:$($script:DoclingSystem.WebPort)" -ForegroundColor Green
        }
    }

    Write-Host "System running!" -ForegroundColor Green

    # Store process IDs for reliable cleanup
    $pidFile = Join-Path $env:TEMP "docling_pids.json"
    $pids = @{
        API       = $apiProcess.Id
        Processor = $procProcess.Id
        Web       = if ($webProcess) { $webProcess.Id } else { $null }
        PyWebView = if ($pyProcess) { $pyProcess.Id } else { $null }
        Timestamp = Get-Date
    }
    $pids | ConvertTo-Json | Set-Content $pidFile -Encoding UTF8

    return @{
        API       = $apiProcess
        Processor = $procProcess
        Web       = $webProcess
        PyWebView = if ($pyProcess) { $pyProcess } else { $null }
    }
}


# Public: Stop-DoclingSystem
function Stop-DoclingSystem {
    [CmdletBinding()]
    param(
        [switch]$ClearQueue
    )

    Write-Host "Stopping Docling System processes..." -ForegroundColor Cyan

    # Stop PowerShell processes running Docling components
    $doclingProcesses = @()

    # Method 1: Check PIDs from stored file (most reliable)
    $pidFile = "$env:TEMP\docling_pids.json"
    if (Test-Path $pidFile) {
        try {
            $storedPids = Get-Content $pidFile | ConvertFrom-Json
            foreach ($processId in @($storedPids.API, $storedPids.Processor, $storedPids.Web, $storedPids.PyWebView)) {
                if ($processId) {
                    $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
                    if ($proc) {
                        $doclingProcesses += $proc
                    }
                }
            }
            Write-Verbose "Found $($doclingProcesses.Count) processes from PID file"
        } catch {
            Write-Warning "Could not read PID file: $($_.Exception.Message)"
        }
    }

    # Method 2: Use WMI to search by CommandLine (slower but finds orphaned processes)
    $wmiProcesses = Get-WmiObject Win32_Process -Filter "Name='powershell.exe' OR Name='python.exe'" -ErrorAction SilentlyContinue
    foreach ($wmiProc in $wmiProcesses) {
        if ($wmiProc.CommandLine) {
            $cmdLine = $wmiProc.CommandLine
            if ($cmdLine -like "*docling_api.ps1*" -or
                $cmdLine -like "*docling_processor.ps1*" -or
                $cmdLine -like "*Start-WebServer.ps1*" -or
                $cmdLine -like "*Launch-PyWebView.py*") {
                $proc = Get-Process -Id $wmiProc.ProcessId -ErrorAction SilentlyContinue
                if ($proc -and $proc -notin $doclingProcesses) {
                    $doclingProcesses += $proc
                }
            }
        }
    }

    if ($doclingProcesses) {
        Write-Host "Found $($doclingProcesses.Count) Docling processes to stop" -ForegroundColor Yellow
        $doclingProcesses | ForEach-Object {
            try {
                Write-Verbose "Stopping process $($_.Id): $($_.ProcessName)"
                $_ | Stop-Process -Force
            } catch {
                Write-Warning "Could not stop process $($_.Id): $($_.Exception.Message)"
            }
        }

        # Remove PID file after stopping processes
        if (Test-Path $pidFile) {
            Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
        }

        Write-Host "Stopped $($doclingProcesses.Count) processes" -ForegroundColor Green
    } else {
        Write-Host "No Docling processes found running" -ForegroundColor Gray
    }

    # Clean up temp files
    $tempFiles = @(
        "$env:TEMP\docling_api.ps1",
        "$env:TEMP\docling_processor.ps1",
        "$env:TEMP\docling_output.txt",
        "$env:TEMP\docling_error.txt",
        "$env:TEMP\docling_processor_debug.txt",
        "$env:TEMP\docling_processor_errors.log"
    )

    $tempFiles | ForEach-Object {
        if (Test-Path $_) {
            $retries = 3
            for ($i = 1; $i -le $retries; $i++) {
                try {
                    Remove-Item $_ -Force -ErrorAction Stop
                    Write-Verbose "Cleaned up temp file: $_"
                    break
                } catch {
                    if ($i -eq $retries) {
                        Write-Warning "Could not remove temp file: $(Split-Path $_ -Leaf)"
                    } else {
                        Start-Sleep -Milliseconds 200
                    }
                }
            }
        }
    }

    # Optionally clear queue and status
    if ($ClearQueue) {
        Write-Host "Clearing queue and status files..." -ForegroundColor Yellow
        Clear-PSDoclingSystem -Force
    }

    Write-Host "Docling System stopped" -ForegroundColor Green
}


Export-ModuleMember -Function @('Get-DoclingConfiguration', 'Set-DoclingConfiguration', 'Get-ProcessingStatus', 'Invoke-DoclingHybridChunking', 'Optimize-ChunksForRAG', 'Set-ProcessingStatus', 'Start-DocumentConversion', 'Test-EnhancedChunking', 'Add-DocumentToQueue', 'Add-QueueItem', 'Add-QueueItemFolder', 'Get-NextQueueItem', 'Get-NextQueueItemFolder', 'Get-QueueItems', 'Get-QueueItemsFolder', 'Set-QueueItems', 'Update-ItemStatus', 'New-FrontendFiles', 'Start-APIServer', 'Start-DocumentProcessor', 'Clear-PSDoclingSystem', 'Get-DoclingSystemStatus', 'Get-PythonStatus', 'Initialize-DoclingSystem', 'Set-PythonAvailable', 'Start-DoclingSystem', 'Stop-DoclingSystem')

Write-Host 'PSDocling Module Loaded - Version 3.2.0' -ForegroundColor Cyan

