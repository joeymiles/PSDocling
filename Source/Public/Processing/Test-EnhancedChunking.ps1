<#
.SYNOPSIS
    Test-EnhancedChunking function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
#>
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
