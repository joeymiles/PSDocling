# PSDocling

PowerShell module that wraps Python [Docling](https://docling-project.github.io/docling/) for document conversion. Convert PDFs, Office files, HTML, Markdown, CSV, and images to Markdown, HTML, JSON, plain text, or DocTags, with an optional REST API and web UI.

**Version:** 3.3.0 · **License:** MIT

## Requirements

- PowerShell 5.1+ or PowerShell Core 6+
- Python 3.8+ with the `docling` package (optional; simulation mode available without Python)
- .NET Framework 4.7.2+ when using Windows PowerShell
- Optional native window: `pip install -r requirements-webview.txt` (PyWebView)

## Install

After clone, run the installer once. It **builds from `Source/` then installs** — no separate Build step.

```powershell
git clone https://github.com/joeymiles/PSDocling.git
cd PSDocling

.\scripts\Install-DoclingModule.ps1
```

Force reinstall:

```powershell
.\scripts\Install-DoclingModule.ps1 -Force
```

## Quick start

```powershell
Import-Module PSDocling
Initialize-DoclingSystem -GenerateFrontend
Start-DoclingSystem -OpenBrowser

# When finished
Stop-DoclingSystem
```

Thin wrappers (same thing from the repo):

```powershell
.\scripts\Start-All.ps1 -GenerateFrontend -OpenBrowser
.\scripts\Stop-All.ps1
```

## Ports and paths

| Component | Default |
|-----------|---------|
| API server | `http://localhost:8080` |
| Web frontend | `http://localhost:8081` |
| Queue folder | `$env:TEMP\DoclingQueue` |
| Status file | `$env:TEMP\docling_status.json` |
| Working temp | `$env:TEMP\DoclingProcessor` |
| Output | `$env:TEMP\DoclingOutput` |

Custom ports via wrapper:

```powershell
.\scripts\Start-All.ps1 -ApiPort 9080 -WebPort 9081 -GenerateFrontend -OpenBrowser
```

Simulation mode (no Python/Docling):

```powershell
.\scripts\Start-All.ps1 -SkipPythonCheck -GenerateFrontend -OpenBrowser
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

- [Overview](Help_files/00_Overview.md)
- [Backend services](Help_files/01_Backend_Services.md)
- [Frontend services](Help_files/02_Frontend_Services.md)
- [File processing](Help_files/03_File_Processing.md)

```powershell
Get-Help Start-DoclingSystem -Full
Get-Help Add-DocumentToQueue -Examples
```

## Repository layout

```
PSDocling/
  Source/           # Module source (authoritative)
  Build/            # Built .psm1/.psd1 (produced by install/build)
  DoclingFrontend/  # Static web UI + Start-WebServer.ps1
  scripts/          # Install, Build, Start/Stop wrappers, PyWebView launcher
  Tests/            # Audit + E2E smoke tests
  Help_files/       # User guides
  PSDocling.psd1    # Manifest (copied into Build/ on build)
```

## Development

```powershell
.\scripts\Build-PSDoclingModule.ps1
.\Tests\Test-AuditFixes.ps1
```

The installer always rebuilds from `Source/` unless you pass `-SkipBuild`. Developers can iterate with Build alone and `Import-Module .\Build\PSDocling.psm1 -Force`.

## Uninstall

```powershell
.\scripts\Uninstall-DoclingModule.ps1
```

## License

MIT — see [LICENSE](LICENSE).
