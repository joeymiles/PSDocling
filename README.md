# PSDocling

PowerShell module that wraps Python [Docling](https://docling-project.github.io/docling/) for document conversion. Convert PDFs, Office files, HTML, Markdown, CSV, and images to Markdown, HTML, JSON, plain text, or DocTags, with an optional REST API and web UI.

**Version:** 3.2.0 · **License:** MIT

## Requirements

- PowerShell 5.1+ or PowerShell Core 6+
- Python 3.8+ with the `docling` package (optional; simulation mode available without Python)
- .NET Framework 4.7.2+ when using Windows PowerShell
- Optional native window: `pip install -r requirements-webview.txt` (PyWebView)

## Install

```powershell
git clone https://github.com/joeymiles/PSDocling.git
cd PSDocling

# Build module from Source/
.\Build-PSDoclingModule.ps1

# Install for current user (or -Scope AllUsers)
.\Install-DoclingModule.ps1
```

To work from the repo without installing:

```powershell
Import-Module .\Build\PSDocling.psm1 -Force
```

## Quick start

```powershell
# Start API + processor + web UI
.\Start-All.ps1 -GenerateFrontend -OpenBrowser

# Stop services
.\Stop-All.ps1
```

Or via module commands:

```powershell
Import-Module PSDocling
Start-DoclingSystem -GenerateFrontend -OpenBrowser
Stop-DoclingSystem
```

## Ports and paths

| Component | Default |
|-----------|---------|
| API server | `http://localhost:8080` |
| Web frontend | `http://localhost:8081` |
| Queue folder | `$env:TEMP\DoclingQueue` |
| Status file | `$env:TEMP\docling_status.json` |
| Working temp | `$env:TEMP\DoclingProcessor` |
| Output | `.\ProcessedDocuments` |

Custom ports:

```powershell
.\Start-All.ps1 -ApiPort 9080 -WebPort 9081 -GenerateFrontend -OpenBrowser
```

Simulation mode (no Python/Docling):

```powershell
.\Start-All.ps1 -SkipPythonCheck -GenerateFrontend -OpenBrowser
```

## Architecture

Three cooperating processes: an HTTP API server, a background document processor that runs Docling, and a static web frontend. Work items move through a folder-based queue with shared status files under `$env:TEMP`.

## Features

- Input: PDF, DOCX, XLSX, PPTX, MD, HTML/XHTML, CSV, PNG/JPEG/TIFF/BMP/WEBP
- Output: Markdown, HTML, JSON, plain text, DocTags (XML)
- Queue-based processing with status tracking
- REST API for programmatic upload/status/download
- Web UI with drag-and-drop upload
- Optional enrichments (code, formulas, picture classification/description)
- Optional hybrid chunking for RAG workflows

## Documentation

Install and usage guides:

- [Overview](Help_files/00_Overview.md)
- [Backend services](Help_files/01_Backend_Services.md)
- [Frontend services](Help_files/02_Frontend_Services.md)
- [File processing](Help_files/03_File_Processing.md)

```powershell
Get-Help Start-DoclingSystem -Full
Get-Help Add-DocumentToQueue -Examples
```

## Development

```powershell
.\Build-PSDoclingModule.ps1
.\Tests\Test-AuditFixes.ps1
```

Source lives under `Source/`; the build concatenates it into `Build/PSDocling.psm1` and copies `PSDocling.psd1`.

## Uninstall

```powershell
.\Uninstall-DoclingModule.ps1
```

## License

MIT — see [LICENSE](LICENSE).
