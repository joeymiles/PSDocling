<#
.SYNOPSIS
    New-FrontendFiles function from PSDocling module
.DESCRIPTION
    Extracted from monolithic PSDocling.psm1
    Original module by: Jaga
.NOTES
    Part of PSDocling Document Processing System
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
