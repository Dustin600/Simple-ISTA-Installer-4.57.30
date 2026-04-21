# ============================================================
# ISTA INSTALLER PATCHES - Drop these into your v3.0 script
# Fixes dual-layout detection (Rheingold vs Modular ISTA)
# ============================================================

# ============================================================
# PATCH 1: Replace Resolve-IstaRoot function
# Old version only checked for TesterGUI as a subfolder.
# New version detects BOTH layouts and returns the config path too.
# ============================================================

function Resolve-IstaRoot {
    param([string]$extractedTo)

    $knownNames = @("ISTA", "ISTA+", "ISTA-PLUS", "ISTA_PLUS", "BMW_ISTA", "ISTAGUI")

    foreach ($name in $knownNames) {
        $candidate = Join-Path $extractedTo $name

        # OLD LAYOUT (Rheingold): has top-level Config\ folder
        if (Test-Path (Join-Path $candidate "Config")) {
            $script:IstaConfigPath   = Join-Path $candidate "Config"
            $script:IstaLayoutStyle  = "rheingold"
            Log-Debug "ISTA root (Rheingold layout) resolved to: $candidate"
            return $candidate
        }

        # NEW LAYOUT (Modular): config lives inside TesterGUI\bin\Release\
        $modularConfig = Join-Path $candidate "TesterGUI\bin\Release"
        if (Test-Path $modularConfig) {
            $script:IstaConfigPath   = $modularConfig
            $script:IstaLayoutStyle  = "modular"
            Log-Debug "ISTA root (Modular layout) resolved to: $candidate"
            return $candidate
        }
    }

    # One level deeper
    $subs = Get-ChildItem -Path $extractedTo -Directory -ErrorAction SilentlyContinue
    foreach ($sub in $subs) {
        if (Test-Path (Join-Path $sub.FullName "Config")) {
            $script:IstaConfigPath   = Join-Path $sub.FullName "Config"
            $script:IstaLayoutStyle  = "rheingold"
            Log-Debug "ISTA root (Rheingold, subfolder) found: $($sub.FullName)"
            return $sub.FullName
        }
        $modularConfig = Join-Path $sub.FullName "TesterGUI\bin\Release"
        if (Test-Path $modularConfig) {
            $script:IstaConfigPath   = $modularConfig
            $script:IstaLayoutStyle  = "modular"
            Log-Debug "ISTA root (Modular, subfolder) found: $($sub.FullName)"
            return $sub.FullName
        }
    }

    # Broader search via ISTAGUI.exe
    $guiSearch = Get-ChildItem "C:\" -Recurse -Filter "ISTAGUI.exe" -ErrorAction SilentlyContinue |
                 Select-Object -First 1
    if ($guiSearch) {
        $root = $guiSearch.Directory.FullName  # start from exe's folder

        # Walk up to find which layout this is
        if (Test-Path (Join-Path ($guiSearch.Directory.Parent.Parent.FullName) "Config")) {
            $script:IstaConfigPath  = Join-Path ($guiSearch.Directory.Parent.Parent.FullName) "Config"
            $script:IstaLayoutStyle = "rheingold"
            return $guiSearch.Directory.Parent.Parent.FullName
        } else {
            # Modular: config is right next to ISTAGUI.exe
            $script:IstaConfigPath  = $guiSearch.DirectoryName
            $script:IstaLayoutStyle = "modular"
            return $guiSearch.Directory.Parent.Parent.FullName
        }
    }

    return $null
}


# ============================================================
# PATCH 2: Add these two script-scope globals near the top of
# the script (with the other $script: globals)
# ============================================================

# Add these lines alongside the existing $script: globals block:
#
#   $script:IstaConfigPath   = $null   # resolved after extraction - layout-aware
#   $script:IstaLayoutStyle  = $null   # "rheingold" or "modular"


# ============================================================
# PATCH 3: Replace the Post-Install Health Check section
# Old version hard-coded "config" folder.
# New version checks the correct path per detected layout.
# ============================================================

Section "Post-Install Health Check"

# Determine config path label for display
$configLabel = if ($script:IstaLayoutStyle -eq "modular") {
    "TesterGUI\bin\Release (config - modular layout)"
} else {
    "Config folder (Rheingold layout)"
}

Write-Host "  Detected layout: $(if ($script:IstaLayoutStyle -eq 'modular') { 'Modular (new)' } else { 'Rheingold (classic)' })" -ForegroundColor Cyan
Write-Host "  Config path    : $($script:IstaConfigPath)" -ForegroundColor DarkGray
Write-Host ""

# Verify the key config files exist at the detected location
$configFiles = @("ISTAGUI.exe.config", "SystemConfig.xml", "DealerData.xml")
$missingConfigs = @()

foreach ($cfgFile in $configFiles) {
    $cfgPath = Join-Path $script:IstaConfigPath $cfgFile
    if (Test-Path $cfgPath) {
        OK "Config file found: $cfgFile"
    } else {
        Write-Host "  [MISSING] $cfgFile not found at: $cfgPath" -ForegroundColor Red
        Log-Info "MISSING config: $cfgFile at $cfgPath"
        $missingConfigs += $cfgFile
    }
}

if ($missingConfigs.Count -gt 0) {
    Write-Host ""
    Write-Host "  [WARN] Missing config files: $($missingConfigs -join ', ')" -ForegroundColor Yellow
    Write-Host "  These are required for ISTA to launch. Your archive may be incomplete." -ForegroundColor Yellow
    Write-Host "  Layout detected: $($script:IstaLayoutStyle)" -ForegroundColor DarkGray
    Write-Host "  Expected config path: $($script:IstaConfigPath)" -ForegroundColor DarkGray
}

$checks = @(
    @{ Path=$script:IstaRoot;          Label="ISTA root folder" },
    @{ Path=$script:IstaExePath;       Label="ISTAGUI.exe" },
    @{ Path=$script:IstaConfigPath;    Label=$configLabel },
    @{ Path="$($script:IstaRoot)\TesterGUI"; Label="TesterGUI folder" },
    @{ Path="$($script:IstaRoot)\SQLiteDBs"; Label="SQLiteDBs folder" },
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

if ($failures -gt 0 -or $missingConfigs.Count -gt 0) {
    Write-Host ""
    Write-Host "  $($failures + $missingConfigs.Count) check(s) failed. Review the items above before launching ISTA." -ForegroundColor Yellow
    Write-Host "  Full log: $logFile" -ForegroundColor Cyan
} else {
    Write-Host ""
    OK "All health checks passed."
}
