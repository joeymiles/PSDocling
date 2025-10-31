# PSDocling Testing Guide

## Ensuring Clean Background Process Management

### The Problem
PSDocling runs three background PowerShell processes:
- API Server (port 8080)
- Document Processor (queue handler)
- Web Server (port 8081)

Without proper cleanup, these processes can become "zombies" that interfere with testing.

### The Solution
The system now has two-tier process tracking:

#### 1. PID File Tracking (Primary Method)
- **File**: `$env:TEMP\docling_pids.json`
- **Created by**: `Start-All.ps1` / `Start-DoclingSystem`
- **Contains**: Process IDs for API, Processor, and Web servers
- **Advantage**: Fast, reliable cleanup

#### 2. WMI Detection (Fallback Method)
- **Uses**: `Get-WmiObject Win32_Process` to search by command line
- **Searches for**: `docling_api.ps1`, `docling_processor.ps1`, `Start-WebServer.ps1`
- **Advantage**: Finds orphaned processes from crashed sessions

### Best Practices for Testing

#### 1. Always Stop Before Starting
```powershell
.\Stop-All.ps1
.\Start-All.ps1 -SkipPythonCheck
```

#### 2. Check for Zombie Processes
```powershell
# Count PowerShell processes
Get-Process powershell,pwsh | Measure-Object | Select-Object -ExpandProperty Count

# See details with WMI
Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" |
    Where-Object { $_.CommandLine -like "*docling*" } |
    Select-Object ProcessId, CommandLine
```

#### 3. Complete Test Cycle
```powershell
# 1. Clean environment
.\Stop-All.ps1

# 2. Start system
.\Start-All.ps1 -SkipPythonCheck

# 3. Run tests
.\test-concurrent.ps1

# 4. Clean shutdown
.\Stop-All.ps1
```

### Diagnostic Scripts

#### test-process-commandline.ps1
Demonstrates why the original Stop-All.ps1 failed:
- Shows `Get-Process` doesn't have CommandLine property
- Demonstrates WMI alternative

#### test-add-queue-debug.ps1
Tests the Add-QueueItem closure fix:
- Verifies items are written to queue
- Shows queue file contents before/after

## Troubleshooting

### "Queue file keeps getting cleared"
**Cause**: Background document processor is running and processing items

**Solution**:
```powershell
.\Stop-All.ps1  # Kills all background processors
```

### "Stop-All.ps1 says 'No processes found' but I see PowerShell processes"
**Cause**: Processes might not be Docling-related

**Check**:
```powershell
# Run the diagnostic
.\test-process-commandline.ps1
```

### "PID file error"
**Cause**: PowerShell's built-in `$PID` variable conflict (fixed in latest version)

**Solution**: Update to latest code (uses `$processId` instead)

## Key Files

- **Stop-All.ps1** - Improved with dual-method detection
- **Start-DoclingSystem.ps1** - Now saves PID file
- **Build/PSDocling.psm1** - Rebuilt module with fixes

## Testing Checklist

- [ ] Run Stop-All.ps1 before each test session
- [ ] Verify no zombie processes after Stop-All.ps1
- [ ] Check PID file is created after Start-All.ps1
- [ ] Confirm all 3 processes start (API, Processor, Web)
- [ ] Verify Stop-All.ps1 finds and kills all 3 processes
- [ ] Confirm PID file is removed after Stop-All.ps1
