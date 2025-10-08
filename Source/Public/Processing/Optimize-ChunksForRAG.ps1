<#
.SYNOPSIS
    Optimize chunking output for RAG (Retrieval Augmented Generation)
.DESCRIPTION
    Post-processes JSONL chunks from Invoke-DoclingHybridChunking to fix issues:
    - Merges split code blocks (unclosed fences)
    - Adds missing metadata (token_count, hash, normalized text)
    - Removes near-duplicate chunks
    - Adjusts chunk sizes to target range (200-400 tokens)
    - Normalizes HTML entities and whitespace
    - Validates and repairs chunk structure
.PARAMETER InputPath
    Path to the JSONL chunks file to optimize
.PARAMETER OutputPath
    Path for the optimized chunks file (defaults to *.optimized.jsonl)
.PARAMETER TargetMinTokens
    Minimum target tokens per chunk (default: 200)
.PARAMETER TargetMaxTokens
    Maximum target tokens per chunk (default: 400)
.PARAMETER MinTokens
    Absolute minimum tokens before merging (default: 50)
.PARAMETER DeduplicationThreshold
    Similarity threshold for near-duplicate detection (0.0-1.0, default: 0.90)
.EXAMPLE
    Optimize-ChunksForRAG -InputPath "document.chunks.jsonl"
.NOTES
    Part of PSDocling Document Processing System
#>
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
    """Estimate token count (4 chars ≈ 1 token)"""
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
