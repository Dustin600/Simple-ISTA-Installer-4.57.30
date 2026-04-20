# ============================================================
# ISTA 4.57.30 UNIVERSAL AUTO-INSTALLER v3.0
# Truly universal - adapts to any system, any file location,
# any naming convention, with full auto-troubleshooting
# ============================================================

param(
    [switch]$DebugMode,
    [switch]$TrimLanguages,
    [switch]$SkipBackup,
    [switch]$ForceReinstall
)

# ============================================================
# TRANSCRIPT / LOG
# ============================================================

$logFile = "$env:USERPROFILE\Desktop\ISTA_Install_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $logFile -Append | Out-Null

# ============================================================
# GLOBALS - populated dynamically at runtime
# ============================================================

$script:ExtractTool     = $null   # path to 7z.exe or rar.exe - found at runtime
$script:ExtractToolType = $null   # "7zip" or "winrar"
$script:IstaRoot        = $null   # resolved after extraction - may not be C:\ISTA
$script:IstaExePath     = $null   # resolved after extraction

# ============================================================
# HELPER: LOGGING
# ============================================================

function Log-Debug {
    param([string]$msg)
    if ($DebugMode) { Write-Host "[DEBUG] $msg" -ForegroundColor DarkGray }
    # Always write to log file regardless of DebugMode
    Add-Content -Path $logFile -Value "[DEBUG] $(Get-Date -Format 'HH:mm:ss') $msg" -ErrorAction SilentlyContinue
}

function Log-Info {
    param([string]$msg)
    Add-Content -Path $logFile -Value "[INFO]  $(Get-Date -Format 'HH:mm:ss') $msg" -ErrorAction SilentlyContinue
}

# ============================================================
# HELPER: UI
# ============================================================

function Section {
    param([string]$title)
    Write-Host ""
    Write-Host "--------------------------------------------" -ForegroundColor DarkCyan
    Write-Host "  $title" -ForegroundColor Cyan
    Write-Host "--------------------------------------------" -ForegroundColor DarkCyan
    Log-Info "=== $title ==="
}

function Warn {
    param([string]$msg)
    Write-Host "  [WARN] $msg" -ForegroundColor Yellow
    Log-Info "WARN: $msg"
}

function OK {
    param([string]$msg)
    Write-Host "  [OK] $msg" -ForegroundColor Green
    Log-Info "OK: $msg"
}

# ============================================================
# HELPER: FAIL WITH AUTO-TROUBLESHOOT
# ============================================================

function Fail {
    param(
        [string]$msg,
        [string]$hint = "",
        [scriptblock]$AutoFix = $null
    )

    Write-Host ""
    Write-Host "  ===============================================" -ForegroundColor Red
    Write-Host "  ERROR: $msg" -ForegroundColor Red
    Write-Host "  ===============================================" -ForegroundColor Red
    Log-Info "FAIL: $msg"

    # Attempt AutoFix if provided
    if ($AutoFix) {
        Write-Host ""
        Write-Host "  [AUTO-FIX] Attempting automatic resolution..." -ForegroundColor Yellow
        try {
            $result = & $AutoFix
            if ($result -eq "FIXED") {
                Write-Host "  [AUTO-FIX] Issue resolved. Continuing..." -ForegroundColor Green
                Log-Info "AutoFix succeeded for: $msg"
                return
            }
        } catch {
            Write-Host "  [AUTO-FIX] Could not resolve automatically: $($_.Exception.Message)" -ForegroundColor Yellow
            Log-Info "AutoFix failed: $($_.Exception.Message)"
        }
    }

    if ($hint) {
        Write-Host ""
        Write-Host "  SUGGESTED FIX:" -ForegroundColor Yellow
        Write-Host "  $hint" -ForegroundColor White
    }

    Write-Host ""
    Write-Host "  Full log: $logFile" -ForegroundColor Cyan
    Stop-Transcript | Out-Null
    Read-Host "  Press ENTER to exit"
    exit 1
}

# ============================================================
# HELPER: FIND EXTRACTION TOOL (7-Zip or WinRAR)
# Searches common install paths + all drives
# ============================================================

function Find-ExtractionTool {
    Section "Locating Extraction Tool (7-Zip or WinRAR)"

    # Known common locations
    $candidates7z = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe",
        "$env:LOCALAPPDATA\Programs\7-Zip\7z.exe",
        "$env:ProgramW6432\7-Zip\7z.exe"
    )

    $candidatesWinRar = @(
        "C:\Program Files\WinRAR\rar.exe",
        "C:\Program Files (x86)\WinRAR\rar.exe",
        "C:\Program Files\WinRAR\WinRAR.exe",
        "C:\Program Files (x86)\WinRAR\WinRAR.exe"
    )

    foreach ($p in $candidates7z) {
        if (Test-Path $p) {
            $script:ExtractTool     = $p
            $script:ExtractToolType = "7zip"
            OK "Found 7-Zip at: $p"
            return
        }
    }

    foreach ($p in $candidatesWinRar) {
        if (Test-Path $p) {
            $script:ExtractTool     = $p
            $script:ExtractToolType = "winrar"
            OK "Found WinRAR at: $p"
            return
        }
    }

    # Not in common paths - scan all drives
    Write-Host "  Not found in common paths. Scanning all drives..." -ForegroundColor Yellow

    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match "^[A-Z]:\\" }

    foreach ($d in $drives) {
        Write-Host "    Scanning $($d.Root) ..." -ForegroundColor DarkGray
        try {
            $found7z = Get-ChildItem $d.Root -Recurse -Filter "7z.exe" -ErrorAction SilentlyContinue |
                       Where-Object { $_.FullName -notmatch "temp|tmp|cache" } |
                       Select-Object -First 1
            if ($found7z) {
                $script:ExtractTool     = $found7z.FullName
                $script:ExtractToolType = "7zip"
                OK "Found 7-Zip at: $($found7z.FullName)"
                return
            }

            $foundRar = Get-ChildItem $d.Root -Recurse -Filter "rar.exe" -ErrorAction SilentlyContinue |
                        Where-Object { $_.FullName -notmatch "temp|tmp|cache" } |
                        Select-Object -First 1
            if ($foundRar) {
                $script:ExtractTool     = $foundRar.FullName
                $script:ExtractToolType = "winrar"
                OK "Found WinRAR at: $($foundRar.FullName)"
                return
            }
        } catch {
            Log-Debug "Scan error on $($d.Root): $($_.Exception.Message)"
        }
    }

    # Last resort - check PATH
    $inPath = Get-Command "7z.exe" -ErrorAction SilentlyContinue
    if ($inPath) {
        $script:ExtractTool     = $inPath.Source
        $script:ExtractToolType = "7zip"
        OK "Found 7-Zip in PATH: $($inPath.Source)"
        return
    }

    Fail "No extraction tool found (7-Zip or WinRAR). Install 7-Zip from https://www.7-zip.org" `
         "Download and install 7-Zip, then re-run this installer." `
         {
            Write-Host "  Attempting to download 7-Zip installer..." -ForegroundColor Yellow
            $dl = "$env:TEMP\7zip_installer.exe"
            try {
                Invoke-WebRequest -Uri "https://www.7-zip.org/a/7z2301-x64.exe" -OutFile $dl -TimeoutSec 30
                if (Test-Path $dl) {
                    Write-Host "  Launching 7-Zip installer - please complete installation then press ENTER." -ForegroundColor Yellow
                    Start-Process $dl -Wait
                    if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
                        $script:ExtractTool     = "C:\Program Files\7-Zip\7z.exe"
                        $script:ExtractToolType = "7zip"
                        return "FIXED"
                    }
                }
            } catch { }
         }
}

# ============================================================
# HELPER: UNIVERSAL EXTRACT
# Handles single RAR, multi-part RAR (.part1.rar / .part01.rar),
# .zip and .7z automatically
# ============================================================

function Invoke-Extract {
    param(
        [string]$archivePath,
        [string]$destination
    )

    if (-not (Test-Path $destination)) {
        New-Item -ItemType Directory -Path $destination -Force | Out-Null
    }

    Log-Debug "Extracting: $archivePath -> $destination using $($script:ExtractToolType)"

    if ($script:ExtractToolType -eq "7zip") {
        & $script:ExtractTool x $archivePath "-o$destination" -y 2>&1 | ForEach-Object {
            if ($_ -match "^(Extracting|Everything)") { Write-Host "    $_" -ForegroundColor DarkGray }
        }
    } elseif ($script:ExtractToolType -eq "winrar") {
        & $script:ExtractTool x -y $archivePath $destination 2>&1 | Out-Null
    }

    return $LASTEXITCODE
}

# ============================================================
# HELPER: UNIVERSAL ARCHIVE FINDER
# - Searches all drives
# - Handles exact names, fuzzy patterns, multi-part sets
# - Shows candidates, lets user pick or enter manually
# - Validates result before returning
# ============================================================

function Find-Archive {
    param(
        [string]$label,
        [string[]]$exactNames,
        [string[]]$fuzzyPatterns,
        [int64]$minBytes,
        [switch]$Optional
    )

    Write-Host ""
    Write-Host "  Searching all drives for $label ..." -ForegroundColor Yellow
    Write-Host "  (Searching: $($exactNames -join ', ') and variations)" -ForegroundColor DarkGray
    Log-Info "Searching for $label | exact=[$($exactNames -join ',')] fuzzy=[$($fuzzyPatterns -join ',')]"

    $allFound = [System.Collections.Generic.List[object]]::new()
    $drives   = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match "^[A-Z]:\\" }

    foreach ($d in $drives) {
        Write-Host "    Scanning $($d.Root) ..." -ForegroundColor DarkGray

        $allPatterns = (@() + $exactNames + $fuzzyPatterns) | Select-Object -Unique

        foreach ($pat in $allPatterns) {
            try {
                $hits = Get-ChildItem -Path $d.Root -Recurse -Filter $pat -ErrorAction SilentlyContinue
                foreach ($h in $hits) {
                    if ($allFound.FullName -notcontains $h.FullName) {
                        $allFound.Add($h)
                        Log-Debug "Candidate: $($h.FullName) ($([math]::Round($h.Length/1MB,0)) MB)"
                    }
                }
            } catch {
                Log-Debug "Scan error [$pat] on $($d.Root): $($_.Exception.Message)"
            }
        }
    }

    # Separate valid from too-small
    $valid    = @($allFound | Where-Object { $_.Length -ge $minBytes } | Sort-Object Length -Descending)
    $tooSmall = @($allFound | Where-Object { $_.Length -lt  $minBytes })

    if ($tooSmall.Count -gt 0) {
        Write-Host ""
        Write-Host "  Files found but below minimum size ($([math]::Round($minBytes/1MB,0)) MB) - skipped:" -ForegroundColor DarkYellow
        foreach ($f in $tooSmall) {
            Write-Host "    [SKIP] $($f.FullName) ($([math]::Round($f.Length/1MB,0)) MB)" -ForegroundColor DarkYellow
        }
        Write-Host ""
        Write-Host "  These files may be incomplete downloads, demo versions, or wrong files." -ForegroundColor DarkYellow
    }

    # ---- Nothing valid found ----
    if ($valid.Count -eq 0) {
        Write-Host ""
        if ($allFound.Count -eq 0) {
            Write-Host "  No files matching '$label' patterns were found on any drive." -ForegroundColor Red
        } else {
            Write-Host "  Files were found but none meet the minimum size requirement." -ForegroundColor Red
        }

        if ($Optional) {
            Write-Host "  This file is optional - skipping." -ForegroundColor Yellow
            Log-Info "$label not found - optional, skipping."
            return $null
        }

        Write-Host ""
        Write-Host "  Options:" -ForegroundColor White
        Write-Host "  [1] Enter the full path manually" -ForegroundColor Cyan
        Write-Host "  [2] Accept a too-small file anyway (not recommended)" -ForegroundColor Cyan
        Write-Host "  [3] Abort installation" -ForegroundColor Red
        Write-Host ""

        $choice = Read-Host "  Enter 1, 2 or 3"

        if ($choice -eq "1") {
            $manual = (Read-Host "  Full path to $label").Trim()
            if ($manual -ne "" -and (Test-Path $manual)) {
                $item = Get-Item $manual
                OK "Using manually specified file: $($item.FullName)"
                Log-Info "Manual path accepted: $($item.FullName)"
                return $item
            }
            Fail "Path not found: $manual"
        } elseif ($choice -eq "2" -and $tooSmall.Count -gt 0) {
            Warn "Accepting undersized file - installation may fail or be incomplete."
            return ($tooSmall | Sort-Object Length -Descending | Select-Object -First 1)
        } else {
            Fail "$label is required and was not found." "Ensure the drive containing your ISTA files is connected and try again."
        }
    }

    # ---- Exactly one valid candidate ----
    if ($valid.Count -eq 1) {
        $f = $valid[0]
        Write-Host ""
        Write-Host "  Found 1 candidate for $label :" -ForegroundColor Green
        Write-Host "    Path     : $($f.FullName)" -ForegroundColor White
        Write-Host "    Size     : $([math]::Round($f.Length/1GB,3)) GB" -ForegroundColor DarkGray
        Write-Host "    Modified : $($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
        Write-Host ""

        $confirm = Read-Host "  Use this file? [Y] Yes  [N] Specify different path  (Y/N)"
        if ($confirm -match "^[Nn]$") {
            $manual = (Read-Host "  Full path to correct $label").Trim()
            if ($manual -ne "" -and (Test-Path $manual)) {
                $item = Get-Item $manual
                OK "Using: $($item.FullName)"
                return $item
            }
            Warn "Invalid path entered - using auto-detected file."
        }

        OK "Confirmed: $($f.FullName)"
        return $f
    }

    # ---- Multiple valid candidates - user picks ----
    Write-Host ""
    Write-Host "  Multiple candidates found for $label - select the correct one:" -ForegroundColor Yellow
    Write-Host ""

    for ($i = 0; $i -lt $valid.Count; $i++) {
        $f = $valid[$i]
        Write-Host "  [$($i+1)] $($f.FullName)" -ForegroundColor White
        Write-Host "      Size: $([math]::Round($f.Length/1GB,3)) GB   Modified: $($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor DarkGray
    }

    $manOpt = $valid.Count + 1
    Write-Host "  [$manOpt] Enter path manually" -ForegroundColor Cyan
    Write-Host ""

    do {
        $raw = Read-Host "  Enter number (1-$manOpt)"
    } while ($raw -notmatch "^\d+$" -or [int]$raw -lt 1 -or [int]$raw -gt $manOpt)

    if ([int]$raw -eq $manOpt) {
        $manual = (Read-Host "  Full path to $label").Trim()
        if ($manual -ne "" -and (Test-Path $manual)) {
            $item = Get-Item $manual
            OK "Using: $($item.FullName)"
            return $item
        }
        Fail "Manual path not found: $manual"
    }

    $selected = $valid[[int]$raw - 1]
    OK "Selected: $($selected.FullName)"
    return $selected
}

# ============================================================
# HELPER: RESOLVE ISTA ROOT DYNAMICALLY
# ISTA may not always extract to C:\ISTA - find it wherever it is
# ============================================================

function Resolve-IstaRoot {
    param([string]$extractedTo)

    # Common folder names ISTA extracts to
    $knownNames = @("ISTA", "ISTA+", "ISTA-PLUS", "ISTA_PLUS", "BMW_ISTA", "ISTAGUI")

    # Check directly under extraction target
    foreach ($name in $knownNames) {
        $candidate = Join-Path $extractedTo $name
        if (Test-Path (Join-Path $candidate "TesterGUI")) {
            Log-Debug "ISTA root resolved to: $candidate"
            return $candidate
        }
    }

    # Search one level deeper
    $subs = Get-ChildItem -Path $extractedTo -Directory -ErrorAction SilentlyContinue
    foreach ($sub in $subs) {
        if (Test-Path (Join-Path $sub.FullName "TesterGUI")) {
            Log-Debug "ISTA root found in subfolder: $($sub.FullName)"
            return $sub.FullName
        }
    }

    # Broader search - look for ISTAGUI.exe on C: drive
    $guiSearch = Get-ChildItem "C:\" -Recurse -Filter "ISTAGUI.exe" -ErrorAction SilentlyContinue |
                 Select-Object -First 1
    if ($guiSearch) {
        $root = $guiSearch.Directory.Parent.Parent.FullName
        Log-Debug "ISTA root found via ISTAGUI.exe search: $root"
        return $root
    }

    return $null
}

# ============================================================
# HELPER: AUTO-TROUBLESHOOT EXTRACTION FAILURE
# ============================================================

function Diagnose-ExtractionFailure {
    param(
        [string]$archivePath,
        [int]$exitCode
    )

    Write-Host ""
    Write-Host "  [TROUBLESHOOT] Diagnosing extraction failure..." -ForegroundColor Yellow

    # Check if file exists
    if (-not (Test-Path $archivePath)) {
        Write-Host "  [DIAG] File no longer accessible: $archivePath" -ForegroundColor Red
        Write-Host "  [DIAG] Was the drive disconnected during extraction?" -ForegroundColor Yellow
        return
    }

    $file = Get-Item $archivePath

    # Check if file is locked
    try {
        $stream = [System.IO.File]::Open($archivePath, 'Open', 'Read', 'None')
        $stream.Close()
        Write-Host "  [DIAG] File is not locked by another process." -ForegroundColor DarkGray
    } catch {
        Write-Host "  [DIAG] File is LOCKED by another process - close any programs using it." -ForegroundColor Red
        return
    }

    # Check free space
    $destDrive = "C"
    $freeGB = [math]::Round((Get-PSDrive $destDrive).Free / 1GB, 1)
    if ($freeGB -lt 5) {
        Write-Host "  [DIAG] CRITICAL: Only ${freeGB} GB free on C: - extraction needs at least 15 GB." -ForegroundColor Red
        return
    }

    # Check exit code meaning
    $exitMeaning = switch ($exitCode) {
        1  { "Warning - non-fatal errors during extraction" }
        2  { "Fatal error - archive may be corrupted" }
        7  { "Command line error" }
        8  { "Not enough memory" }
        255{ "User stopped the process" }
        default { "Unknown error (code $exitCode)" }
    }
    Write-Host "  [DIAG] Exit code $exitCode means: $exitMeaning" -ForegroundColor Yellow

    # Check if multi-part archive
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($archivePath)
    $dir      = [System.IO.Path]::GetDirectoryName($archivePath)
    $parts    = Get-ChildItem -Path $dir -Filter "$baseName.part*.rar" -ErrorAction SilentlyContinue
    if ($parts.Count -gt 1) {
        $missing = @()
        for ($i = 1; $i -le $parts.Count; $i++) {
            $partName = "$baseName.part$($i.ToString('D2')).rar"
            if (-not (Test-Path (Join-Path $dir $partName))) {
                $missing += $partName
            }
        }
        if ($missing.Count -gt 0) {
            Write-Host "  [DIAG] Multi-part archive missing parts: $($missing -join ', ')" -ForegroundColor Red
        } else {
            Write-Host "  [DIAG] All $($parts.Count) archive parts are present." -ForegroundColor DarkGray
        }
    }

    # Antivirus check
    $avProcesses = @("MsMpEng","avguard","avgnt","ekrn","bdagent","mbam","nod32","avg","avp","mcshield")
    $runningAV = Get-Process | Where-Object { $avProcesses -contains $_.Name.ToLower() }
    if ($runningAV) {
        Write-Host "  [DIAG] Antivirus detected: $($runningAV.Name -join ', ')" -ForegroundColor Yellow
        Write-Host "  [DIAG] Try temporarily disabling AV and re-running - AV often blocks ISTA archives." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "  [DIAG] Suggestions:" -ForegroundColor White
    Write-Host "    1. Verify the archive is not corrupted (re-download if needed)" -ForegroundColor White
    Write-Host "    2. Temporarily disable antivirus software" -ForegroundColor White
    Write-Host "    3. Ensure destination drive has 15+ GB free" -ForegroundColor White
    Write-Host "    4. Run script from a local drive, not a network share" -ForegroundColor White
}

# ============================================================
# HELPER: SAFE BACKUP
# ============================================================

function Backup-IfExists {
    param([string]$path, [string]$label)
    if (Test-Path $path) {
        $stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backup = "${path}_backup_$stamp"
        Write-Host "  Existing $label found. Backing up..." -ForegroundColor Yellow
        try {
            Rename-Item -Path $path -NewName $backup -ErrorAction Stop
            OK "$label backed up to: $backup"
            Log-Info "Backed up $path to $backup"
        } catch {
            Warn "Could not rename $path - trying copy instead."
            Copy-Item -Path $path -Destination $backup -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path $backup) {
                Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                OK "Backup via copy/delete completed: $backup"
            } else {
                Fail "Cannot back up existing installation at $path. Close any programs using ISTA and retry." `
                     "Make sure ISTA is not running and no files in $path are open."
            }
        }
    } else {
        Write-Host "  [OK] No existing $label found. Skipping backup." -ForegroundColor Green
    }
}

# ============================================================
# BANNER
# ============================================================

Clear-Host
Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "   BMW ISTA-PLUS 4.57.30 Universal Installer" -ForegroundColor White
Write-Host "            v3.0 - Any System Edition" -ForegroundColor DarkGray
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "  Debug Mode   : $DebugMode" -ForegroundColor DarkGray
Write-Host "  TrimLanguages: $TrimLanguages" -ForegroundColor DarkGray
Write-Host "  SkipBackup   : $SkipBackup" -ForegroundColor DarkGray
Write-Host "  Log File     : $logFile" -ForegroundColor DarkGray
Write-Host ""

# ============================================================
# SECTION 1 - ADMIN CHECK
# ============================================================

Section "Checking Administrator Privileges"

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Fail "Must run as Administrator." "Right-click the .bat launcher and choose Run as Administrator." {
        # Attempt self-elevate
        $psi = New-Object System.Diagnostics.ProcessStartInfo "powershell"
        $psi.Arguments  = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        $psi.Verb        = "runas"
        $psi.WindowStyle = "Normal"
        [System.Diagnostics.Process]::Start($psi) | Out-Null
        exit
    }
}
OK "Running as Administrator."

# ============================================================
# SECTION 2 - EXECUTION POLICY
# ============================================================

Section "Checking PowerShell Execution Policy"

$policy = Get-ExecutionPolicy
Log-Debug "Execution policy: $policy"

if ($policy -eq "Restricted") {
    Write-Host "  Execution policy is Restricted - attempting auto-fix..." -ForegroundColor Yellow
    try {
        Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
        OK "Execution policy updated to RemoteSigned."
    } catch {
        Fail "Cannot set execution policy: $($_.Exception.Message)" `
             "Run: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
    }
} else {
    OK "Execution policy: $policy"
}

# ============================================================
# SECTION 3 - WINDOWS VERSION
# ============================================================

Section "Checking Windows Version"

$winVer   = [System.Environment]::OSVersion.Version
$winBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
$winName  = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName

Log-Debug "Windows: $winName Build $winBuild ($($winVer.ToString()))"

if ($winVer.Major -lt 10) {
    Fail "Windows 10 or 11 (64-bit) is required. Detected: $winName"
}

OK "Windows: $winName (Build $winBuild)"

# ============================================================
# SECTION 4 - .NET FRAMEWORK
# ============================================================

Section "Checking .NET Framework"

$dotnet = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction SilentlyContinue

if (-not $dotnet -or $dotnet.Release -lt 461808) {
    $releaseFound = if ($dotnet) { $dotnet.Release } else { "not found" }
    Fail ".NET 4.7.2+ required. Found release key: $releaseFound" `
         "Download .NET 4.7.2 from: https://dotnet.microsoft.com/download/dotnet-framework" `
         {
            Write-Host "  Attempting .NET 4.8 download..." -ForegroundColor Yellow
            try {
                $dlPath = "$env:TEMP\ndp48.exe"
                Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?LinkId=2085155" -OutFile $dlPath -TimeoutSec 60
                Write-Host "  Installing .NET 4.8 - please complete the installer then press ENTER." -ForegroundColor Yellow
                Start-Process $dlPath -Wait
                return "FIXED"
            } catch {
                Write-Host "  Auto-download failed: $($_.Exception.Message)" -ForegroundColor Yellow
            }
         }
}

OK ".NET Framework release key: $($dotnet.Release) (4.7.2+ confirmed)"

# ============================================================
# SECTION 5 - FIND EXTRACTION TOOL
# ============================================================

Find-ExtractionTool

# ============================================================
# SECTION 6 - DISK SPACE
# ============================================================

Section "Checking Disk Space on C:"

$freeBytes = (Get-PSDrive C).Free
$freeGB    = [math]::Round($freeBytes / 1GB, 1)
Log-Debug "Free on C: $freeGB GB"

if ($freeBytes -lt 15GB) {
    Write-Host ""
    Write-Host "  Only $freeGB GB free on C: - 15 GB minimum required." -ForegroundColor Red
    Write-Host ""
    Write-Host "  [TROUBLESHOOT] Largest folders on C:\" -ForegroundColor Yellow
    try {
        Get-ChildItem "C:\" -Directory -ErrorAction SilentlyContinue |
        ForEach-Object {
            $size = (Get-ChildItem $_.FullName -Recurse -ErrorAction SilentlyContinue |
                     Measure-Object -Property Length -Sum).Sum
            [PSCustomObject]@{ Path=$_.FullName; GB=[math]::Round($size/1GB,1) }
        } | Sort-Object GB -Descending | Select-Object -First 5 |
        ForEach-Object { Write-Host "    $($_.GB) GB  $($_.Path)" -ForegroundColor DarkGray }
    } catch {}
    Fail "Insufficient disk space: $freeGB GB available, 15 GB required." `
         "Free up space on C: or install to a different drive."
}

OK "Free disk space: $freeGB GB on C:"

# ============================================================
# SECTION 7 - LOCATE SOURCE ARCHIVES
# ============================================================

Section "Locating Source Archives"

# ISTA.rar - also handles ISTA.part1.rar, ISTA.part01.rar etc.
$istaRar = Find-Archive `
    -label         "ISTA.rar (main application)" `
    -exactNames    @("ISTA.rar","ISTA.part1.rar","ISTA.part01.rar","ISTA.part001.rar") `
    -fuzzyPatterns @("ISTA*.rar","*ISTA-PLUS*.rar","*ista*.rar","*ISTAGUI*.rar","BMW_ISTA*.rar") `
    -minBytes      500MB

# BMW.rar - ISPI data package
$bmwRar = Find-Archive `
    -label         "BMW.rar (ISPI data)" `
    -exactNames    @("BMW.rar","BMW.part1.rar","BMW.part01.rar") `
    -fuzzyPatterns @("*BMW*.rar","*bmw*.rar","*ISPI*.rar","BMW_data*.rar","*ispi_data*.rar") `
    -minBytes      100MB

# PSDZdata - optional, only needed for full programming
$psdzdataRar = Find-Archive `
    -label         "PSDZdata.rar (ECU programming data)" `
    -exactNames    @("PSDZdata.rar","PSDZdata.part1.rar","PSDZdata.part01.rar") `
    -fuzzyPatterns @("PSDZdata*.rar","*psdzdata*.rar","*PSDZDATA*.rar","*PSDZData*.rar","*ecu_data*.rar") `
    -minBytes      100MB `
    -Optional

# ============================================================
# SECTION 8 - OPERATION MODE
# ============================================================

Section "Select Operation Mode"

Write-Host ""
Write-Host "  [1] Diagnostics Only  - Fault codes, live data, service functions" -ForegroundColor Cyan
Write-Host "  [2] Full Programming  - Diagnostics + ECU coding and flashing" -ForegroundColor Cyan
Write-Host ""

do { $modeChoice = Read-Host "  Enter 1 or 2" } while ($modeChoice -ne "1" -and $modeChoice -ne "2")

$operationMode = if ($modeChoice -eq "1") { "diagnostic" } else { "full" }
OK "Mode: $(if ($operationMode -eq 'diagnostic') { 'Diagnostics Only' } else { 'Full Programming' })"

if ($operationMode -eq "full" -and -not $psdzdataRar) {
    Write-Host ""
    Warn "Full Programming selected but PSDZdata was not found."
    Warn "ECU programming will NOT work. Download PSDZdata separately if needed."
    Write-Host ""
    Read-Host "  Press ENTER to continue without PSDZdata, or Ctrl+C to cancel"
}

# ============================================================
# SECTION 9 - KILL RUNNING ISTA
# ============================================================

Section "Stopping Any Running ISTA Processes"

$istaProcs = @("ISTAGUI","TesterGUI","ISTAClient","ISTAService","Rheingold")
foreach ($pName in $istaProcs) {
    $running = Get-Process -Name $pName -ErrorAction SilentlyContinue
    if ($running) {
        Write-Host "  Stopping: $pName ..." -ForegroundColor Yellow
        $running | Stop-Process -Force
        Start-Sleep -Seconds 1
        OK "Stopped: $pName"
    }
}
OK "No ISTA processes running."

# ============================================================
# SECTION 10 - BACKUP
# ============================================================

if (-not $SkipBackup) {
    Section "Backing Up Existing Installations"
    Backup-IfExists "C:\ISTA"             "C:\ISTA"
    Backup-IfExists "C:\ProgramData\BMW"  "C:\ProgramData\BMW"
} else {
    Warn "SkipBackup flag set - skipping backup."
}

# ============================================================
# SECTION 11 - EXTRACT ISTA.RAR
# ============================================================

Section "Step 1: Extracting ISTA Package"

Write-Host "  Source : $($istaRar.FullName)" -ForegroundColor DarkGray
Write-Host "  Target : C:\" -ForegroundColor DarkGray
Write-Host "  Extracting - this may take several minutes..." -ForegroundColor Yellow

$exitCode = Invoke-Extract -archivePath $istaRar.FullName -destination "C:\"

if ($exitCode -ne 0 -and $exitCode -ne 1) {
    Diagnose-ExtractionFailure -archivePath $istaRar.FullName -exitCode $exitCode
    Fail "ISTA.rar extraction failed (exit code: $exitCode). See diagnostics above."
}

# Dynamically locate where ISTA actually extracted
$script:IstaRoot = Resolve-IstaRoot -extractedTo "C:\"

if (-not $script:IstaRoot) {
    Write-Host ""
    Write-Host "  ISTA folder not found at expected location. Searching C: drive..." -ForegroundColor Yellow
    $guiExe = Get-ChildItem "C:\" -Recurse -Filter "ISTAGUI.exe" -ErrorAction SilentlyContinue |
              Select-Object -First 1
    if ($guiExe) {
        $script:IstaRoot = $guiExe.DirectoryName | Split-Path | Split-Path
        Write-Host "  Found ISTA at: $($script:IstaRoot)" -ForegroundColor Green
    } else {
        Fail "ISTAGUI.exe not found anywhere on C: after extraction." `
             "ISTA.rar may be corrupted or may require a different extraction path."
    }
}

$script:IstaExePath = Get-ChildItem -Path $script:IstaRoot -Recurse -Filter "ISTAGUI.exe" -ErrorAction SilentlyContinue |
                      Select-Object -First 1 | Select-Object -ExpandProperty FullName

OK "ISTA extracted to: $($script:IstaRoot)"
OK "ISTAGUI.exe found at: $($script:IstaExePath)"

# If ISTA extracted to a non-standard folder, offer to create symlink at C:\ISTA
if ($script:IstaRoot -ne "C:\ISTA" -and -not (Test-Path "C:\ISTA")) {
    Write-Host "  ISTA is at '$($script:IstaRoot)' rather than C:\ISTA." -ForegroundColor Yellow
    Write-Host "  Creating junction (C:\ISTA -> $($script:IstaRoot)) for compatibility..." -ForegroundColor Yellow
    cmd /c mklink /J "C:\ISTA" $script:IstaRoot | Out-Null
    if (Test-Path "C:\ISTA") {
        OK "Junction created: C:\ISTA -> $($script:IstaRoot)"
    } else {
        Warn "Could not create junction. Using $($script:IstaRoot) directly."
    }
}

# ============================================================
# SECTION 12 - FIND AND EXTRACT BMW.RAR
# ============================================================

Section "Step 3: Extracting BMW.rar (ISPI Data)"

# If not found in initial scan, search inside the now-extracted ISTA folder
if (-not $bmwRar) {
    Write-Host "  Searching for BMW.rar inside extracted ISTA package..." -ForegroundColor Yellow
    $bmwRar = Get-ChildItem -Path $script:IstaRoot -Recurse -Filter "BMW*.rar" -ErrorAction SilentlyContinue |
              Where-Object { $_.Length -gt 100MB } |
              Sort-Object Length -Descending |
              Select-Object -First 1

    if ($bmwRar) {
        OK "Found BMW.rar inside ISTA package: $($bmwRar.FullName)"
    } else {
        Write-Host ""
        Write-Host "  BMW.rar still not found. Enter path manually or press ENTER to abort." -ForegroundColor Red
        $manual = (Read-Host "  Full path to BMW.rar").Trim()
        if ($manual -ne "" -and (Test-Path $manual)) {
            $bmwRar = Get-Item $manual
            OK "Using manually specified BMW.rar: $($bmwRar.FullName)"
        } else {
            Fail "BMW.rar is required and could not be found." `
                 "Ensure the drive with your ISTA files is connected. BMW.rar should be ~150-300 MB."
        }
    }
}

Write-Host "  Source : $($bmwRar.FullName)" -ForegroundColor DarkGray
Write-Host "  Target : C:\ProgramData" -ForegroundColor DarkGray
Write-Host "  Extracting..." -ForegroundColor Yellow

$exitCode = Invoke-Extract -archivePath $bmwRar.FullName -destination "C:\ProgramData"

if ($exitCode -ne 0 -and $exitCode -ne 1) {
    Diagnose-ExtractionFailure -archivePath $bmwRar.FullName -exitCode $exitCode
    Fail "BMW.rar extraction failed (exit code: $exitCode)."
}

OK "BMW.rar extracted to C:\ProgramData"

# ============================================================
# SECTION 13 - ENVIRONMENT VARIABLES
# ============================================================

Section "Step 4: Setting ISPI Environment Variables"

$ispiBase = "C:\ProgramData\BMW\ISPI"

# Try to auto-locate ISPI if not at expected path
if (-not (Test-Path $ispiBase)) {
    Write-Host "  ISPI not at expected path. Searching C:\ProgramData\BMW\ ..." -ForegroundColor Yellow
    $ispiSearch = Get-ChildItem "C:\ProgramData\BMW" -Recurse -Filter "ISPI" -Directory -ErrorAction SilentlyContinue |
                  Select-Object -First 1
    if ($ispiSearch) {
        $ispiBase = $ispiSearch.FullName
        Warn "ISPI found at non-standard path: $ispiBase"
    }
}

$envVars = @{
    "ISPI_DIR"  = $ispiBase
    "ISPI_DATA" = "$ispiBase\data"
    "ISPI_LOG"  = "$ispiBase\logs"
}

foreach ($key in $envVars.Keys) {
    [Environment]::SetEnvironmentVariable($key, $envVars[$key], "Machine")
    OK "$key = $($envVars[$key])"
}

# Create log/data dirs if missing
foreach ($dir in @("$ispiBase\data","$ispiBase\logs")) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        OK "Created missing directory: $dir"
    }
}

# ============================================================
# SECTION 14 - REGISTRY
# ============================================================

Section "Step 5: Writing Registry Settings"

$regPath = "HKLM\SOFTWARE\BMWGroup\ISPI\ISTA"

try {
    if ($operationMode -eq "diagnostic") {
        reg add $regPath /v "OperationMode"      /t REG_SZ    /d "diagnostic" /f | Out-Null
        reg add $regPath /v "ProgrammingEnabled" /t REG_DWORD /d 0             /f | Out-Null
        OK "Registry: Diagnostics Only mode."
    } else {
        reg add $regPath /v "OperationMode"      /t REG_SZ    /d "full" /f | Out-Null
        reg add $regPath /v "ProgrammingEnabled" /t REG_DWORD /d 1       /f | Out-Null
        OK "Registry: Full Programming mode."
    }

    # Write ISTA root path to registry for apps that need it
    reg add $regPath /v "InstallPath" /t REG_SZ /d $script:IstaRoot /f | Out-Null
    OK "Registry: InstallPath = $($script:IstaRoot)"
} catch {
    Fail "Registry write failed: $($_.Exception.Message)" "Ensure you are running as Administrator."
}

# ============================================================
# SECTION 15 - PSDZDATA
# ============================================================

if ($operationMode -eq "full") {
    Section "Step 6: PSDZdata Setup (ECU Programming)"

    $psdzdataTarget = "$($script:IstaRoot)\PdZ\data_swi"

    if ($psdzdataRar) {
        if (-not (Test-Path $psdzdataTarget)) {
            New-Item -ItemType Directory -Path $psdzdataTarget -Force | Out-Null
        }

        Write-Host "  Source : $($psdzdataRar.FullName)" -ForegroundColor DarkGray
        Write-Host "  Target : $psdzdataTarget" -ForegroundColor DarkGray
        Write-Host "  Extracting PSDZdata - this can take 20+ minutes..." -ForegroundColor Yellow

        $exitCode = Invoke-Extract -archivePath $psdzdataRar.FullName -destination $psdzdataTarget

        if ($exitCode -ne 0 -and $exitCode -ne 1) {
            Diagnose-ExtractionFailure -archivePath $psdzdataRar.FullName -exitCode $exitCode
            Warn "PSDZdata extraction reported errors. ECU programming may not work correctly."
        } else {
            OK "PSDZdata extracted to: $psdzdataTarget"
        }
    } else {
        Warn "PSDZdata not available - ECU programming disabled."
        Warn "Download from: https://binunlock.com/resources/psdzdata-full-ecu-programming-data-pack-for-e-sys.229/"
        Warn "Extract manually to: $($script:IstaRoot)\PdZ\data_swi"
    }
}

# ============================================================
# SECTION 16 - D-CAN SETUP
# ============================================================

Section "Step 7: D-CAN Interface Setup"

$dcanDir  = Join-Path $script:IstaRoot "tool"
$dcanTool = Get-ChildItem -Path $dcanDir -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "DCAN|Switch" } |
            Select-Object -First 1

if ($dcanTool) {
    Write-Host "  D-CAN tool found: $($dcanTool.Name)" -ForegroundColor White
    Write-Host "  Are you using a D-CAN (OBD) cable? [Y/N]" -ForegroundColor White
    $dcanChoice = Read-Host "  Y or N"

    if ($dcanChoice -match "^[Yy]$") {
        Read-Host "  Press ENTER to launch D-CAN tool (set COM3, Latency=1 when it opens)"
        Start-Process $dcanTool.FullName -Wait -Verb RunAs
        OK "D-CAN tool completed."
    } else {
        OK "D-CAN skipped - using ICOM/ENET."
    }
} else {
    Warn "D-CAN tool not found in $dcanDir - skip if using ICOM/ENET."
}

# ============================================================
# SECTION 17 - LANGUAGE TRIM (OPTIONAL)
# ============================================================

if ($TrimLanguages) {
    Section "Trimming Non-English SQLite Language Files"
    $sqlitePath = Join-Path $script:IstaRoot "SQLiteDBs"
    if (Test-Path $sqlitePath) {
        $removed = 0
        Get-ChildItem $sqlitePath -Recurse -Filter "*.sqlite" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch "EN\.sqlite" } |
        ForEach-Object { Remove-Item $_.FullName -Force; $removed++ }
        OK "Removed $removed non-EN SQLite files."
    } else {
        Warn "SQLiteDBs folder not found at $sqlitePath - skipping trim."
    }
}

# ============================================================
# SECTION 18 - DESKTOP SHORTCUT
# ============================================================

Section "Step 2: Creating Desktop Shortcut"

if (Test-Path $script:IstaExePath) {
    $shortcut = "$env:Public\Desktop\ISTA 4.57.30.lnk"
    try {
        $wsh              = New-Object -ComObject WScript.Shell
        $s                = $wsh.CreateShortcut($shortcut)
        $s.TargetPath       = $script:IstaExePath
        $s.IconLocation     = "$($script:IstaExePath), 0"
        $s.WorkingDirectory = Split-Path $script:IstaExePath
        $s.Description      = "BMW ISTA-PLUS 4.57.30"
        $s.Save()
        OK "Shortcut created: $shortcut"
    } catch {
        Warn "Could not create shortcut: $($_.Exception.Message)"
        Warn "You can manually launch: $($script:IstaExePath)"
    }
} else {
    Warn "ISTAGUI.exe not found - shortcut skipped."
}

# ============================================================
# SECTION 19 - POST-INSTALL HEALTH CHECK
# ============================================================

Section "Post-Install Health Check"

$checks = @(
    @{ Path=$script:IstaRoot;          Label="ISTA root folder" },
    @{ Path=$script:IstaExePath;       Label="ISTAGUI.exe" },
    @{ Path="$($script:IstaRoot)\TesterGUI"; Label="TesterGUI folder" },
    @{ Path="$($script:IstaRoot)\SQLiteDBs"; Label="SQLiteDBs folder" },
    @{ Path="$($script:IstaRoot)\config";    Label="config folder" },
    @{ Path="$ispiBase";               Label="ISPI base folder" },
    @{ Path="$ispiBase\data";          Label="ISPI data folder" }
)

$failures = 0
foreach ($chk in $checks) {
    if (Test-Path $chk.Path) {
        OK "$($chk.Label): $($chk.Path)"
    } else {
        Write-Host "  [MISSING] $($chk.Label): $($chk.Path)" -ForegroundColor Red
        Log-Info "MISSING: $($chk.Label) at $($chk.Path)"
        $failures++
    }
}

# SQLite EN DB count
$sqlitePath = Join-Path $script:IstaRoot "SQLiteDBs"
if (Test-Path $sqlitePath) {
    $enDbs = @(Get-ChildItem $sqlitePath -Recurse -Filter "*EN*.sqlite" -ErrorAction SilentlyContinue)
    if ($enDbs.Count -lt 3) {
        Write-Host "  [WARN] Only $($enDbs.Count) EN SQLite DBs found (expected 3+)" -ForegroundColor Yellow
        $failures++
    } else {
        OK "$($enDbs.Count) EN SQLite DB files confirmed."
    }
}

if ($failures -gt 0) {
    Write-Host ""
    Write-Host "  $failures check(s) failed. Review the items above before launching ISTA." -ForegroundColor Yellow
    Write-Host "  Full log: $logFile" -ForegroundColor Cyan
} else {
    Write-Host ""
    OK "All health checks passed."
}

# ============================================================
# DONE
# ============================================================

Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "   BMW ISTA-PLUS 4.57.30 Install Complete!" -ForegroundColor White
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Mode      : $(if ($operationMode -eq 'diagnostic') { 'Diagnostics Only' } else { 'Full Programming' })" -ForegroundColor White
Write-Host "  ISTA Root : $($script:IstaRoot)" -ForegroundColor White
Write-Host "  Shortcut  : Desktop > ISTA 4.57.30" -ForegroundColor White
Write-Host "  Log file  : $logFile" -ForegroundColor White
Write-Host ""
Write-Host "  IMPORTANT: Always launch ISTA as Administrator." -ForegroundColor Yellow

if ($operationMode -eq "full") {
    Write-Host ""
    Write-Host "  WARNING: Full Programming mode active." -ForegroundColor Red
    Write-Host "           Exercise extreme caution when coding/flashing ECUs." -ForegroundColor Red
}

Write-Host ""
Stop-Transcript | Out-Null

Read-Host "Press ENTER to launch ISTA now"
if (Test-Path $script:IstaExePath) {
    Start-Process $script:IstaExePath -Verb RunAs
} else {
    Write-Host "  Could not find ISTAGUI.exe to launch automatically." -ForegroundColor Yellow
    Write-Host "  Navigate to $($script:IstaRoot) and launch ISTAGUI.exe manually." -ForegroundColor White
}
