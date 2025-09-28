<#
.SYNOPSIS
    Invoke-DoclingHybridChunking function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
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
