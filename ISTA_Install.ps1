# ============================================================
# ISTA 4.57.30 UNIVERSAL AUTO-INSTALLER
# Based on official BMW ISTA-PLUS 4.57.30 installation guide
# Supports: Diagnostic Only or Full Programming mode
# ============================================================

param(
    [switch]$DebugMode,
    [switch]$TrimLanguages
)

# ============================================================
# TRANSCRIPT / LOG FILE
# ============================================================

$logFile = "$env:USERPROFILE\Desktop\ISTA_Install_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $logFile -Append | Out-Null

# ============================================================
# HELPER FUNCTIONS
# ============================================================

function Log-Debug {
    param([string]$msg)
    if ($DebugMode) {
        Write-Host "[DEBUG] $msg" -ForegroundColor DarkGray
    }
}

function Fail {
    param([string]$msg)
    Write-Host ""
    Write-Host "===============================================" -ForegroundColor Red
    Write-Host "  ERROR: $msg" -ForegroundColor Red
    Write-Host "===============================================" -ForegroundColor Red
    Write-Host "  Install log saved to: $logFile" -ForegroundColor Cyan
    Stop-Transcript | Out-Null
    Read-Host "Press ENTER to exit"
    exit 1
}

function Section {
    param([string]$title)
    Write-Host ""
    Write-Host "--------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "  $title" -ForegroundColor Cyan
    Write-Host "--------------------------------------------" -ForegroundColor DarkCyan
}

function Warn {
    param([string]$msg)
    Write-Host "  [WARN] $msg" -ForegroundColor Yellow
    Log-Debug "WARNING: $msg"
}

function Find-Archive {
    param(
        [string]$fileNamePattern,
        [int64]$minBytes
    )

    Write-Host "  Searching all fixed drives for $fileNamePattern ..." -ForegroundColor Yellow
    Log-Debug "Searching all drives for $fileNamePattern (min size: $([math]::Round($minBytes/1MB,1)) MB)"

    $candidates = @()

    $drives = Get-PSDrive -PSProvider FileSystem |
              Where-Object { $_.Free -gt 0 -and $_.Root -match "^[A-Z]:\\" }

    foreach ($d in $drives) {
        Write-Host "    Scanning $($d.Root) ..." -ForegroundColor DarkGray
        Log-Debug "Scanning drive $($d.Root)"

        try {
            $found = Get-ChildItem -Path $d.Root -Recurse -Filter $fileNamePattern -ErrorAction SilentlyContinue |
                     Where-Object { $_.Length -ge $minBytes }

            if ($found) {
                $candidates += $found
                foreach ($f in $found) {
                    Write-Host "      Found: $($f.FullName) ($([math]::Round($f.Length/1GB,2)) GB)" -ForegroundColor DarkGreen
                    Log-Debug "Candidate: $($f.FullName)"
                }
            }
        } catch {
            Log-Debug "Access denied or error on $($d.Root): $($_.Exception.Message)"
        }
    }

    if ($candidates.Count -gt 0) {
        $best = $candidates | Sort-Object Length -Descending | Select-Object -First 1
        Write-Host "  [OK] Using: $($best.FullName)" -ForegroundColor Green
        Write-Host "       Size : $([math]::Round($best.Length / 1GB, 2)) GB" -ForegroundColor DarkGray
        Log-Debug "Selected archive: $($best.FullName)"
        return $best
    }

    Log-Debug "No valid archive found for $fileNamePattern"
    return $null
}

# ============================================================
# FIXED PATHS (per official install guide)
# ============================================================

$istaRoot        = "C:\ISTA"
$istaConfigDir   = "C:\ISTA+4.57.30\config"
$istaExePath     = "C:\ISTA\TesterGUI\bin\Release\ISTAGUI.exe"
$psdzdataTarget  = "C:\ISTA\PdZ\data_swi"
$dcanToolDir     = "C:\ISTA\tool"
$ispiBase        = "C:\ProgramData\BMW\ISPI"
$sevenZip        = "C:\Program Files\7-Zip\7z.exe"

# Minimum file sizes
$minIstaRarBytes = 500MB
$minBmwRarBytes  = 100MB

# ============================================================
# BANNER
# ============================================================

Clear-Host
Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "   BMW ISTA-PLUS 4.57.30 Auto-Installer" -ForegroundColor White
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "  Debug Mode   : $DebugMode" -ForegroundColor DarkGray
Write-Host "  TrimLanguages: $TrimLanguages (keep EN only)" -ForegroundColor DarkGray
Write-Host "  Log File     : $logFile" -ForegroundColor DarkGray
Write-Host ""

# ============================================================
# SECTION 1 - ADMIN CHECK
# ============================================================

Section "Checking Administrator Privileges"

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Fail "This script must be run as Administrator. Right-click and select Run as Administrator."
}

Write-Host "  [OK] Running as Administrator." -ForegroundColor Green
Log-Debug "Admin check passed."

# ============================================================
# SECTION 2 - EXECUTION POLICY CHECK
# ============================================================

Section "Checking PowerShell Execution Policy"

$policy = Get-ExecutionPolicy
Log-Debug "Current execution policy: $policy"

if ($policy -eq "Restricted") {
    Fail "PowerShell execution policy is Restricted. Fix: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
}

Write-Host "  [OK] Execution policy is: $policy" -ForegroundColor Green

# ============================================================
# SECTION 3 - WINDOWS VERSION CHECK
# ============================================================

Section "Checking Windows Version"

$winVer = [System.Environment]::OSVersion.Version
Log-Debug "Detected Windows version: $($winVer.ToString())"

if ($winVer.Major -lt 10) {
    Fail "Windows 10/11 x64 is required. Detected: $($winVer.ToString())"
}

Write-Host "  [OK] Windows version: $($winVer.ToString())" -ForegroundColor Green

# ============================================================
# SECTION 4 - .NET FRAMEWORK CHECK
# ============================================================

Section "Checking .NET Framework"

$dotnet = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue

if (-not $dotnet -or $dotnet.Release -lt 461808) {
    Fail ".NET Framework 4.7.2 or later is required. Install .NET 4.7.2+ and re-run."
}

Log-Debug ".NET release key: $($dotnet.Release)"
Write-Host "  [OK] .NET Framework release: $($dotnet.Release) (4.7.2+ confirmed)" -ForegroundColor Green

# ============================================================
# SECTION 5 - 7-ZIP CHECK
# ============================================================

Section "Checking 7-Zip"

if (-not (Test-Path $sevenZip)) {
    Fail "7-Zip not found at: $sevenZip - Install from https://www.7-zip.org before running."
}

Log-Debug "7-Zip found at $sevenZip"
Write-Host "  [OK] 7-Zip found." -ForegroundColor Green

# ============================================================
# SECTION 6 - DISK SPACE CHECK
# ============================================================

Section "Checking Disk Space"

$freeSpace = (Get-PSDrive C).Free
$freeGB    = [math]::Round($freeSpace / 1GB, 1)
Log-Debug "Free space on C: $freeGB GB"

if ($freeSpace -lt 15GB) {
    Fail "Not enough disk space. 15GB required, only $freeGB GB available on C:"
}

Write-Host "  [OK] Free disk space: $freeGB GB" -ForegroundColor Green

# ============================================================
# SECTION 7 - LOCATE SOURCE ARCHIVES (FULL ALL-DRIVE SCAN)
# ============================================================

Section "Locating Source Archives (Scanning All Drives)"

# --- ISTA.rar ---
$istaRar = Find-Archive -fileNamePattern "ISTA.rar" -minBytes $minIstaRarBytes
if (-not $istaRar) {
    Fail "ISTA.rar not found on any drive. Minimum size: 500MB. Ensure the drive is connected and accessible."
}

# --- BMW.rar ---
$bmwRar = Find-Archive -fileNamePattern "BMW.rar" -minBytes $minBmwRarBytes
if ($bmwRar) {
    Log-Debug "BMW.rar found: $($bmwRar.FullName) ($([math]::Round($bmwRar.Length/1GB,2)) GB)"
    Write-Host "  [OK] Found BMW.rar: $($bmwRar.FullName)" -ForegroundColor Green
    Write-Host "       Size: $([math]::Round($bmwRar.Length / 1GB, 2)) GB" -ForegroundColor DarkGray
} else {
    Warn "BMW.rar not found on any drive. Will search again inside extracted ISTA package."
}

# --- PSDZdata (optional) ---
$psdzdataRar = Find-Archive -fileNamePattern "PSDZdata*.rar" -minBytes 100MB
if ($psdzdataRar) {
    Write-Host "  [OK] Found PSDZdata: $($psdzdataRar.FullName)" -ForegroundColor Green
    Log-Debug "PSDZdata RAR: $($psdzdataRar.FullName)"
} else {
    Warn "PSDZdata archive not found on any drive. ECU programming will not work without it."
}

# ============================================================
# SECTION 8 - OPERATION MODE SELECTION
# ============================================================

Section "Select Operation Mode"

Write-Host ""
Write-Host "  Choose ISTA operation mode:" -ForegroundColor White
Write-Host ""
Write-Host "  [1] Diagnostics Only  - Read fault codes, live data, service functions" -ForegroundColor Cyan
Write-Host "  [2] Full Programming  - Diagnostics + ECU coding and programming" -ForegroundColor Cyan
Write-Host ""

do {
    $modeChoice = Read-Host "  Enter 1 or 2"
} while ($modeChoice -ne "1" -and $modeChoice -ne "2")

if ($modeChoice -eq "1") {
    $operationMode = "diagnostic"
    Write-Host ""
    Write-Host "  [OK] Mode selected: Diagnostics Only" -ForegroundColor Green
    Log-Debug "Operation mode: Diagnostics Only"
} else {
    $operationMode = "full"
    Write-Host ""
    Write-Host "  [OK] Mode selected: Full Programming" -ForegroundColor Green
    Log-Debug "Operation mode: Full Programming"

    if (-not $psdzdataRar) {
        Write-Host ""
        Warn "Full Programming selected but PSDZdata was not found on any drive."
        Warn "ECU programming will NOT work without PSDZdata."
        Warn "Download from: https://binunlock.com/resources/psdzdata-full-ecu-programming-data-pack-for-e-sys.229/"
        Warn "Then extract manually to: $psdzdataTarget"
        Write-Host ""
        Read-Host "  Press ENTER to continue anyway, or Ctrl+C to cancel"
    }
}

# ============================================================
# SECTION 9 - KILL RUNNING ISTA
# ============================================================

Section "Checking for Running ISTA Processes"

$running = Get-Process -Name "ISTAGUI" -ErrorAction SilentlyContinue

if ($running) {
    Write-Host "  ISTA is currently running. Closing before install..." -ForegroundColor Yellow
    $running | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "  [OK] ISTA process closed." -ForegroundColor Green
    Log-Debug "ISTAGUI process terminated."
} else {
    Write-Host "  [OK] No running ISTA processes found." -ForegroundColor Green
    Log-Debug "No ISTAGUI process running."
}

# ============================================================
# SECTION 10 - BACKUP EXISTING INSTALLATIONS
# ============================================================

Section "Backing Up Existing Installations"

if (Test-Path $istaRoot) {
    $istaBackup = "C:\ISTA_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Write-Host "  Existing C:\ISTA found. Backing up..." -ForegroundColor Yellow
    Rename-Item -Path $istaRoot -NewName $istaBackup -ErrorAction Stop
    Write-Host "  [OK] C:\ISTA backed up to: $istaBackup" -ForegroundColor Green
    Log-Debug "C:\ISTA backed up to $istaBackup"
} else {
    Write-Host "  [OK] No existing C:\ISTA found. Skipping." -ForegroundColor Green
}

if (Test-Path "C:\ProgramData\BMW") {
    $bmwBackup = "C:\ProgramData\BMW_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Write-Host "  Existing ProgramData\BMW found. Backing up..." -ForegroundColor Yellow
    Rename-Item -Path "C:\ProgramData\BMW" -NewName $bmwBackup -ErrorAction Stop
    Write-Host "  [OK] ProgramData\BMW backed up to: $bmwBackup" -ForegroundColor Green
    Log-Debug "ProgramData\BMW backed up to $bmwBackup"
} else {
    Write-Host "  [OK] No existing ProgramData\BMW found. Skipping." -ForegroundColor Green
}

# ============================================================
# SECTION 11 - STEP 1: EXTRACT ISTA.RAR TO C:\ISTA
# ============================================================

Section "Step 1: Extracting ISTA.rar to C:\ISTA"

Write-Host "  Extracting ISTA.rar - this may take several minutes..." -ForegroundColor Yellow
Log-Debug "Running: $sevenZip x $($istaRar.FullName) -oC:\ -y"

& $sevenZip x $istaRar.FullName "-oC:\" -y

if ($LASTEXITCODE -ne 0) {
    Fail "7-Zip extraction of ISTA.rar failed (exit code: $LASTEXITCODE). File may be corrupted."
}

if (-not (Test-Path $istaExePath)) {
    Fail "ISTAGUI.exe not found at $istaExePath after extraction. ISTA.rar may extract to a different folder name."
}

Write-Host "  [OK] ISTA.rar extracted to C:\ISTA" -ForegroundColor Green
Write-Host "  [OK] ISTAGUI.exe confirmed at: $istaExePath" -ForegroundColor Green
Log-Debug "ISTA.rar extraction complete."

# ============================================================
# SECTION 12 - STEP 3: FIND AND EXTRACT BMW.RAR
# ============================================================

Section "Step 3: Extracting BMW.rar to C:\ProgramData"

# If BMW.rar was not found during initial scan, search again inside extracted ISTA folder
if (-not $bmwRar) {
    Write-Host "  Searching for BMW.rar inside extracted ISTA folder..." -ForegroundColor Yellow

    $configBmwRar = Join-Path $istaConfigDir "BMW.rar"
    if (Test-Path $configBmwRar) {
        $item = Get-Item $configBmwRar
        if ($item.Length -gt $minBmwRarBytes) {
            $bmwRar = $item
            Log-Debug "BMW.rar found at config path after extraction: $configBmwRar"
        }
    }

    if (-not $bmwRar) {
        $bmwRar = Get-ChildItem -Path $istaRoot -Recurse -Filter "BMW.rar" -ErrorAction SilentlyContinue |
                  Where-Object { $_.Length -gt $minBmwRarBytes } |
                  Sort-Object Length -Descending |
                  Select-Object -First 1
    }

    if (-not $bmwRar) {
        Fail "BMW.rar not found on any drive or inside the extracted ISTA package. Ensure the drive holding your files is connected."
    }

    Write-Host "  [OK] Found BMW.rar after extraction: $($bmwRar.FullName)" -ForegroundColor Green
}

$bmwRarFinal = $bmwRar.FullName
Log-Debug "Using BMW.rar: $bmwRarFinal ($([math]::Round($bmwRar.Length/1GB,2)) GB)"

Write-Host "  Extracting BMW.rar - this may take several minutes..." -ForegroundColor Yellow

& $sevenZip x $bmwRarFinal "-oC:\ProgramData" -y

if ($LASTEXITCODE -ne 0) {
    Fail "7-Zip extraction of BMW.rar failed (exit code: $LASTEXITCODE). File may be corrupted."
}

Write-Host "  [OK] BMW.rar extracted to C:\ProgramData" -ForegroundColor Green
Log-Debug "BMW.rar extraction complete."

# ============================================================
# SECTION 13 - STEP 4: SET ISPI SYSTEM VARIABLES
# ============================================================

Section "Step 4: Setting ISPI System Variables"

$envVars = @{
    "ISPI_DIR"  = "C:\ProgramData\BMW\ISPI"
    "ISPI_DATA" = "C:\ProgramData\BMW\ISPI\data"
    "ISPI_LOG"  = "C:\ProgramData\BMW\ISPI\logs"
}

foreach ($key in $envVars.Keys) {
    [Environment]::SetEnvironmentVariable($key, $envVars[$key], "Machine")
    Log-Debug "Set $key = $($envVars[$key])"
    Write-Host "  [OK] $key = $($envVars[$key])" -ForegroundColor Green
}

# ============================================================
# SECTION 14 - STEP 5: REGISTRY OPERATION MODE
# ============================================================

Section "Step 5: Writing Registry Operation Mode"

if ($operationMode -eq "diagnostic") {
    reg add "HKLM\SOFTWARE\BMWGroup\ISPI\ISTA" /v "OperationMode"      /t REG_SZ    /d "diagnostic" /f | Out-Null
    reg add "HKLM\SOFTWARE\BMWGroup\ISPI\ISTA" /v "ProgrammingEnabled" /t REG_DWORD /d 0             /f | Out-Null
    Write-Host "  [OK] Registry set: Diagnostics Only mode." -ForegroundColor Green
    Log-Debug "Registry: diagnostic mode applied."
} else {
    reg add "HKLM\SOFTWARE\BMWGroup\ISPI\ISTA" /v "OperationMode"      /t REG_SZ    /d "full" /f | Out-Null
    reg add "HKLM\SOFTWARE\BMWGroup\ISPI\ISTA" /v "ProgrammingEnabled" /t REG_DWORD /d 1       /f | Out-Null
    Write-Host "  [OK] Registry set: Full Programming mode." -ForegroundColor Green
    Log-Debug "Registry: full programming mode applied."
}

if ($LASTEXITCODE -ne 0) {
    Fail "Registry write failed (exit code: $LASTEXITCODE). Ensure you are running as Administrator."
}

# ============================================================
# SECTION 15 - STEP 6: PSDZDATA (FULL PROGRAMMING ONLY)
# ============================================================

if ($operationMode -eq "full") {
    Section "Step 6: PSDZdata Setup (Full Programming)"

    if ($psdzdataRar) {
        if (-not (Test-Path $psdzdataTarget)) {
            New-Item -ItemType Directory -Path $psdzdataTarget -Force | Out-Null
            Log-Debug "Created PSDZdata target: $psdzdataTarget"
        }

        Write-Host "  Extracting PSDZdata - this may take 20+ minutes..." -ForegroundColor Yellow
        Log-Debug "Running: $sevenZip x $($psdzdataRar.FullName) -o$psdzdataTarget -y"

        & $sevenZip x $psdzdataRar.FullName "-o$psdzdataTarget" -y

        if ($LASTEXITCODE -ne 0) {
            Fail "PSDZdata extraction failed (exit code: $LASTEXITCODE)."
        }

        Write-Host "  [OK] PSDZdata extracted to: $psdzdataTarget" -ForegroundColor Green
        Log-Debug "PSDZdata extraction complete."
    } else {
        Warn "PSDZdata not found - skipping. Extract manually to: $psdzdataTarget"
        Warn "Download: https://binunlock.com/resources/psdzdata-full-ecu-programming-data-pack-for-e-sys.229/"
    }
} else {
    Log-Debug "Diagnostic mode - skipping PSDZdata."
}

# ============================================================
# SECTION 16 - STEP 7: D-CAN INTERFACE SETUP
# ============================================================

Section "Step 7: D-CAN Interface Setup"

$dcanTool = Get-ChildItem -Path $dcanToolDir -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "DCAN|Switch" } |
            Select-Object -First 1

if ($dcanTool) {
    Write-Host ""
    Write-Host "  D-CAN switch tool found: $($dcanTool.Name)" -ForegroundColor White
    Write-Host ""
    Write-Host "  Are you using a D-CAN (OBD) cable to connect to your BMW?" -ForegroundColor White
    Write-Host "  [Y] Yes - configure D-CAN (COM3, Latency: 1)" -ForegroundColor Cyan
    Write-Host "  [N] No  - using ICOM/ENET, skip this step" -ForegroundColor Cyan
    Write-Host ""

    $dcanChoice = Read-Host "  Enter Y or N"

    if ($dcanChoice -match "^[Yy]$") {
        Write-Host ""
        Write-Host "  Launching D-CAN switch tool..." -ForegroundColor Yellow
        Write-Host "  When it opens: press any key, set Interface=COM3 and Latency Timer=1" -ForegroundColor Yellow
        Write-Host ""
        Read-Host "  Press ENTER to launch the tool now"
        Start-Process $dcanTool.FullName -Wait -Verb RunAs
        Write-Host "  [OK] D-CAN tool launched and completed." -ForegroundColor Green
        Log-Debug "D-CAN switch tool ran: $($dcanTool.FullName)"
    } else {
        Write-Host "  [OK] D-CAN skipped. ICOM/ENET mode retained." -ForegroundColor Green
        Log-Debug "D-CAN setup skipped by user."
    }
} else {
    Warn "D-CAN switch tool not found in $dcanToolDir"
    Warn "If using a D-CAN cable, manually run the switch tool from C:\ISTA\tool\ before launching ISTA."
}

# ============================================================
# SECTION 17 - OPTIONAL SQLITE LANGUAGE TRIM
# ============================================================

if ($TrimLanguages) {
    Section "Trimming SQLite Languages (Keeping EN Only)"

    $sqlitePath = Join-Path $istaRoot "SQLiteDBs"

    if (-not (Test-Path $sqlitePath)) {
        Warn "SQLiteDBs folder not found. Skipping trim."
    } else {
        $dbFiles = Get-ChildItem -Path $sqlitePath -Recurse -Filter "*.sqlite" -ErrorAction SilentlyContinue
        $removed = 0
        foreach ($db in $dbFiles) {
            if ($db.Name -notmatch "EN\.sqlite") {
                Remove-Item $db.FullName -Force
                $removed++
                Log-Debug "Removed: $($db.FullName)"
            }
        }
        Write-Host "  [OK] Language trim complete. Removed $removed non-EN DB file(s)." -ForegroundColor Green
    }
} else {
    Log-Debug "TrimLanguages not set; skipping."
}

# ============================================================
# SECTION 18 - SQLITE SANITY CHECK
# ============================================================

Section "SQLiteDB Sanity Check"

$sqlitePath = Join-Path $istaRoot "SQLiteDBs"

if (-not (Test-Path $sqlitePath)) {
    Fail "SQLiteDBs folder missing - ISTA will not function correctly."
}

$enDbs = Get-ChildItem -Path $sqlitePath -Recurse -Filter "*EN*.sqlite" -ErrorAction SilentlyContinue

if (-not $enDbs -or $enDbs.Count -lt 3) {
    Fail "Insufficient EN SQLite DBs found ($($enDbs.Count)). Install may be incomplete."
}

Write-Host "  [OK] Found $($enDbs.Count) EN SQLite DB file(s)." -ForegroundColor Green
Log-Debug "EN SQLite DB count: $($enDbs.Count)"

# ============================================================
# SECTION 19 - STEP 2: CREATE DESKTOP SHORTCUT
# ============================================================

Section "Step 2: Creating Desktop Shortcut"

if (-not (Test-Path $istaExePath)) {
    Fail "ISTAGUI.exe not found at: $istaExePath"
}

$shortcut   = "$env:Public\Desktop\ISTA 4.57.30.lnk"
$WshShell   = New-Object -ComObject WScript.Shell
$s          = $WshShell.CreateShortcut($shortcut)
$s.TargetPath       = $istaExePath
$s.IconLocation     = "$istaExePath, 0"
$s.WorkingDirectory = (Split-Path $istaExePath)
$s.Description      = "BMW ISTA-PLUS 4.57.30 Diagnostic Tool"
$s.Save()

Write-Host "  [OK] Desktop shortcut created: ISTA 4.57.30" -ForegroundColor Green
Log-Debug "Shortcut created at $shortcut"

# ============================================================
# SECTION 20 - POST-INSTALL HEALTH CHECK
# ============================================================

Section "Post-Install Health Check"

if (-not (Test-Path $istaRoot))    { Fail "C:\ISTA missing after extraction." }
Write-Host "  [OK] C:\ISTA exists." -ForegroundColor Green

if (-not (Test-Path $istaExePath)) { Fail "ISTAGUI.exe missing at $istaExePath" }
Write-Host "  [OK] ISTAGUI.exe confirmed." -ForegroundColor Green

foreach ($folder in @("TesterGUI", "SQLiteDBs", "config", "tool")) {
    if (-not (Test-Path "$istaRoot\$folder")) { Fail "Missing ISTA folder: $folder" }
    Write-Host "  [OK] C:\ISTA\$folder exists." -ForegroundColor Green
    Log-Debug "ISTA folder OK: $folder"
}

if (-not (Test-Path $ispiBase)) { Fail "C:\ProgramData\BMW\ISPI missing." }
Write-Host "  [OK] C:\ProgramData\BMW\ISPI exists." -ForegroundColor Green

foreach ($folder in @("data", "logs", "ISTA", "TRIC")) {
    if (-not (Test-Path "$ispiBase\$folder")) { Fail "Missing ISPI component: $folder" }
    Write-Host "  [OK] ISPI\$folder exists." -ForegroundColor Green
    Log-Debug "ISPI folder OK: $folder"
}

if ($operationMode -eq "full" -and $psdzdataRar) {
    if (-not (Test-Path $psdzdataTarget)) {
        Warn "PSDZdata folder not confirmed at $psdzdataTarget - ECU programming may not work."
    } else {
        Write-Host "  [OK] PSDZdata exists at: $psdzdataTarget" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "  Post-install health check passed." -ForegroundColor Green

# ============================================================
# DONE
# ============================================================

Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "   BMW ISTA-PLUS 4.57.30 Install Complete!" -ForegroundColor White
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mode     : $(if ($operationMode -eq 'diagnostic') { 'Diagnostics Only' } else { 'Full Programming' })" -ForegroundColor White
Write-Host "  Shortcut : Desktop > ISTA 4.57.30" -ForegroundColor White
Write-Host "  Log file : $logFile" -ForegroundColor White
Write-Host ""
Write-Host "  IMPORTANT: Always launch ISTA as Administrator." -ForegroundColor Yellow

if ($operationMode -eq "full") {
    Write-Host ""
    Write-Host "  WARNING: Full Programming mode is active." -ForegroundColor Red
    Write-Host "           Exercise extreme caution when coding or flashing ECUs." -ForegroundColor Red
}

Write-Host ""

Stop-Transcript | Out-Null

Read-Host "Press ENTER to launch ISTA now"
Start-Process $istaExePath -Verb RunAs
