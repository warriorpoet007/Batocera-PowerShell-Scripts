<#
PURPOSE: Create .m3u files for each multi-disk game and insert the list of game filenames into the playlist or update gamelist.xml
VERSION: 1.3
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Place this file into the ROMS folder to process all platforms, or in a platform's individual subfolder to process just that one.
- False detections and misses are possible, especially for complex naming structures, but should be rare.
    - Built-in intelligence attempts to determine and annotate why:
        - A playlist might have been suppressed
        - A multi-disk file wasn't incorporated into a playlist
- There are three variables you can update to change script behavior:
    - $dryRun : default is $false; set to $true to preview output without modifying files
    - $nonM3UPlatforms : array of platforms (folder names) that should NOT use M3U; default: 3DO + apple2
    - $noM3UPlatformMode : default "move"; set to "skip" to ignore NON-M3U platforms entirely

BREAKDOWN
- Enumerates ROM game files starting in the directory the script resides in
    - Scans up to 2 subdirectory levels deep recursively
    - Skips .m3u files during scanning (so it doesn’t treat playlists as input)
    - Skips common media/manual folders (e.g., images, videos, media, manuals, downloaded_*) to reduce false multi-disk detections
    - For platforms that can't use M3U playlist files, the script instead hides Disk 2+ in gamelist.xml (<hidden>true</hidden>)
        - Per run, a backup of gamelist.xml is made first called gamelist.backup, labeled with (1), (2), etc. if a backup already exists.
        - This creates a single entry in Batocera for a multi-disk game (for the first disk in the set) instead of one for each disk
        - Initially, this includes 3DO and Apple II but additional platform folders can be added into the $nonM3UPlatforms array
        - If you'd rather just skip these platforms, change the $noM3UPlatformMode variable to "skip" instead of "move"
- Detects multi-disk candidates by parsing filenames for “designators”
    - A designator is a disk/disc/side marker that indicates a set (case-insensitive), such as:
        - Disk 1, Disc B, Disk II, Disk 2 of 6, Disk 4 Side A, etc.
        - Side-only sets like Side A, Side B, etc. are supported (treated as Disk 1 with different sides)
    - Supports disk tokens as:
        - Numbers (1, 2, …)
        - Letters (A, B, …)
        - Roman numerals (I … XX) (used for sort normalization)
    - Also recognizes optional patterns like:
        - of N totals (e.g., Disk 2 of 6)
        - Side X paired with a disk marker (e.g., Disk 2 Side B)
- Extracts and interprets bracket tags for grouping and playlist naming
    - Separates tags into:
        - Alt tags like [a], [a2], [b], [b3] (TOSEC-style), etc.
        - Other base tags like [cr ...], [! ], etc.
    - Uses a “non-bang” compatibility key:
        - Treats sets as compatible when only [!] differs across files (helpful when some disks include [!] and others don’t)
- Groups files into candidate multi-disk sets and selects the best disk entries
    - Primary grouping is strict by:
        - directory + base game title prefix + base tags (excluding alt tags)
    - Uses multiple passes to fill disk slots robustly:
        - strict match within the group
        - relaxed match within the same title where only [!] differs
        - alt fallback chain support (e.g., [a2] → [a] → base) when matching is incomplete
        - conservative “base playlist can accept a single unambiguous alt disk” rule (to avoid missing a disk when only one variant exists)
- Builds a stable playlist filename
    - Playlist filename is based on:
        - the normalized base title prefix
        - optional shared name hint (only if all selected entries share it)
        - only tags common across all selected entries (prevents one-off tags from polluting the name)
        - appends the normalized alt tag only if all selected entries share the same alt
    - Cleans up the final playlist name (removes double spaces, dangling punctuation/parens, etc.)
- Prevents collisions and duplicates
    - Same-run path collisions:
        - If the intended playlist path is already “claimed” in the current run (written or suppressed), it generates alternate names:
            - [alt], [alt2], etc.
    - Same-run duplicate playlist content suppression:
        - If an identical ordered list of disk files would be emitted again during the same run, it suppresses the duplicate and reports what it duplicated.
- Writes .m3u playlists with strict cleanliness rules
    - Writes the filenames only (not full paths), in disk/side order
    - Ensures playlists written by the script have:
        - no trailing whitespace at end of any line
        - no blank lines
        - no trailing newline at EOF
        - written as UTF-8 without BOM
    - If an existing .m3u already exists:
        - If content is identical (after normalizing newline style and BOM only), it is suppressed (not overwritten)
        - If content is different (including extra blank lines), it is overwritten and flagged as such in the report
    - Or uses similar logic to find a NON-M3U game's disk entries after the first in the set in gamelist.xml and tag them as hidden
- Tracks which disk files were “used”
    - Files included in either written playlists or suppressed playlists are marked “used”
    - For NON-M3U platforms, Disk 2+ entries successfully hidden in gamelist.xml are marked “used”
    - Remaining parsed multi-disk candidates that weren’t used are reported as:
        - (POSSIBLE) MULTI-DISK FILES SKIPPED, with a reason such as:
        - incomplete disk set
        - missing matching disk
        - suppressed by [!] preference rule
        - alt fallback issues
        - disk total mismatch issues
        - no entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)
- Reporting and summary output
    - Clean split into:
        - M3U PLAYLISTS (created + suppressed)
        - GAMELIST(S) UPDATED (NON-M3U platforms)
    - Displays runtime as:
        - "X seconds" (<60s)
        - "M:SS" (<60m)
        - "H:MM:SS" (>=60m)
#>

# ==================================================================================================
# SCRIPT STARTUP: PATHS, TIMING, AND COUNTERS
# ==================================================================================================

# Establish script working directory (where scanning begins) and start time (for runtime reporting)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptStart = Get-Date

# Establish per-platform playlist counts and a total playlist counter
$platformCounts = @{}
$totalPlaylistsCreated = 0

# --------------------------------------------------------------------------------------------------
# USER CONFIGURATION: DRY RUN AND NON-M3U PLATFORM HANDLING
# --------------------------------------------------------------------------------------------------

# Establish DRY RUN mode:
# - $true  => do not write .m3u files and do not modify gamelist.xml; report what WOULD happen
# - $false => perform writes/edits normally
$dryRun = $false # <-- set to $true if you want to see what the output would be without changing files

# Establish platforms that should NOT use .m3u playlists (use ROM folder names for systems)
$nonM3UPlatforms = @(
    '3DO'
    'apple2'
)

# Establish how NON-M3U platforms are handled:
# - "move" => identify sets and hide Disk 2+ in gamelist.xml
# - "skip" => ignore these platforms entirely (no M3U and no gamelist edits)
$noM3UPlatformMode = "move"   # <-- set to "skip" to completely ignore those platforms

# --------------------------------------------------------------------------------------------------
# CONSOLE OUTPUT SAFETY: BUFFER WIDTH
# --------------------------------------------------------------------------------------------------

# Attempt to widen console buffer to reduce truncation of long paths in output (best-effort)
try {
    $raw = $Host.UI.RawUI
    $size = $raw.BufferSize
    if ($size.Width -lt 300) {
        $raw.BufferSize = New-Object Management.Automation.Host.Size(250, $size.Height)
    }
} catch {
    # Ignore if host doesn't allow resizing (e.g., some terminals)
}

# --------------------------------------------------------------------------------------------------
# SCANNING FILTERS: FOLDERS TO SKIP
# --------------------------------------------------------------------------------------------------

# Establish folder names that should not be scanned (reduces false disk detections)
$skipFolders = @(
    'images','videos','media','manuals',
    'downloaded_images','downloaded_videos','downloaded_media','downloaded_manuals'
)

# --------------------------------------------------------------------------------------------------
# GAMELIST STATE + REPORTING BUCKETS (NON-M3U WORKFLOW)
# --------------------------------------------------------------------------------------------------

# Establish cached gamelist state per platform (prevents repeated ReadAllLines calls)
$gamelistStateByPlatform = @{}     # platformLower -> state object (cached lines)

# Establish per-run gamelist backup tracking (one backup per gamelist.xml per run)
$gamelistBackupDone      = @{}     # gamelist path -> $true

# Establish a quick lookup for "missing from gamelist.xml" (used for skip reason selection)
$noM3UMissingGamelistEntryByFullPath = @{}  # full file path -> $true

# Use ArrayList buckets to avoid op_Addition crashes in large runs
$noM3UPrimaryEntriesOk         = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UPrimaryEntriesIncomplete = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

$noM3UNewlyHidden              = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UAlreadyHidden            = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UMissingGamelistEntries   = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

# Establish per-platform count buckets for gamelist changes
$gamelistHiddenCounts        = @{}  # platform label -> count newly hidden
$gamelistAlreadyHiddenCounts = @{}  # platform label -> count already hidden (no change)

# ==================================================================================================
# FUNCTIONS
# ==================================================================================================

# --- FUNCTION: Normalize-M3UText ---
# PURPOSE:
# - Normalize M3U text for stable equality checks by ignoring BOM and newline style differences.
# NOTES:
# - Preserves trailing spaces and trailing blank lines as meaningful differences.
# - Returns an array of lines (including trailing empty lines).
function Normalize-M3UText {
    param([AllowNull()][string]$Text)

    # If there's no text, treat as empty list of lines
    if ($null -eq $Text) { return @() }

    # Strip UTF-8 BOM if present (only for equality checks)
    if ($Text.Length -gt 0 -and [int]$Text[0] -eq 0xFEFF) {
        $Text = $Text.Substring(1)
    }

    # Normalize newline style only (CRLF/CR => LF)
    $Text = $Text -replace "`r`n", "`n"
    $Text = $Text -replace "`r", "`n"

    # Split into lines, preserving trailing empty lines
    return ,($Text -split "`n")
}

# --- FUNCTION: Convert-DiskToSort ---
# PURPOSE:
# - Convert a disk token (numeric, letter, roman numeral) into a sortable integer.
# NOTES:
# - Returns $null if the token is missing or unrecognized.
function Convert-DiskToSort {
    param([string]$DiskToken)

    # Guard: empty token means no disk info
    if ([string]::IsNullOrWhiteSpace($DiskToken)) { return $null }

    # Numeric disk token
    if ($DiskToken -match '^\d+$') { return [int]$DiskToken }

    # Roman numeral disk token (1..20) — used only for sort normalization
    $romanMap = @{
        'I' = 1;  'II' = 2;  'III' = 3;  'IV' = 4;  'V' = 5
        'VI' = 6; 'VII' = 7; 'VIII' = 8; 'IX' = 9;  'X' = 10
        'XI' = 11; 'XII' = 12; 'XIII' = 13; 'XIV' = 14; 'XV' = 15
        'XVI' = 16; 'XVII' = 17; 'XVIII' = 18; 'XIX' = 19; 'XX' = 20
    }

    $upper = $DiskToken.ToUpperInvariant()
    if ($romanMap.ContainsKey($upper)) { return $romanMap[$upper] }

    # Single-letter disk token: A=1, B=2, ...
    if ($upper -match '^[A-Z]$') {
        return ([int][char]$upper[0]) - 64
    }

    # Unknown token type
    return $null
}

# --- FUNCTION: Convert-SideToSort ---
# PURPOSE:
# - Convert a Side token (A/B/...) into a sortable integer (A=1, B=2, ...).
# NOTES:
# - Returns 0 when side is missing/blank so “no side” sorts first.
function Convert-SideToSort {
    param([string]$SideToken)

    # Guard: no side sorts first
    if ([string]::IsNullOrWhiteSpace($SideToken)) { return 0 }

    $c = $SideToken.ToUpperInvariant()[0]
    return ([int][char]$c) - 64
}

# --- FUNCTION: Is-AltTag ---
# PURPOSE:
# - Detect TOSEC-style alt tags like [a], [a2], [b], [b3].
# NOTES:
# - Used to separate “alt tags” from other bracket tags during parsing.
function Is-AltTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)[ab]\d*\]$')
}

# --- FUNCTION: Is-DiskNoiseTag ---
# PURPOSE:
# - Detect bracketed disk descriptor tags that should not influence grouping (noise).
# NOTES:
# - Prevents tags like “[Disk A]” from changing grouping/playlist names.
function Is-DiskNoiseTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)\s*disks?\b')
}

# --- FUNCTION: Clean-BasePrefix ---
# PURPOSE:
# - Clean/normalize the filename prefix before the disk designator for stable grouping & naming.
# NOTES:
# - Trims trailing punctuation/whitespace and dangling parentheses artifacts.
function Clean-BasePrefix {
    param([string]$Prefix)

    # Guard: null prefix becomes empty string
    if ($null -eq $Prefix) { return "" }

    $p = $Prefix.Trim()

    # Remove common trailing separators/periods/underscores/dashes
    $p = $p -replace '[\s._-]+$', ''

    # Remove dangling "(" at end (artifact from split patterns)
    $p = $p -replace '\(\s*$', ''

    return $p.Trim()
}

# --- FUNCTION: Write-Phase ---
# PURPOSE:
# - Print a short “phase banner” so long operations don’t look like the script is hanging.
# NOTES:
# - Pure console output helper (no behavioral impact).
function Write-Phase {
    param([string]$Text)
    Write-Host ""
    Write-Host $Text -ForegroundColor Cyan
}

# --- FUNCTION: Get-AltFallbackChain ---
# PURPOSE:
# - Build an ordered alt fallback chain (e.g., [a2] -> [a] -> base) used during disk filling.
# NOTES:
# - Always ends with "" (base/no-alt) as the final fallback.
function Get-AltFallbackChain {
    param([string]$AltKey)

    # If there is no alt, only consider base
    if ([string]::IsNullOrWhiteSpace($AltKey)) { return @("") }

    # Parse [a2] / [b] patterns
    $m = [regex]::Match($AltKey, '^\[(?i)(?<L>[ab])(?<N>\d*)\]$')
    if (-not $m.Success) { return @($AltKey, "") }

    $letter = $m.Groups['L'].Value.ToLowerInvariant()
    $num = $m.Groups['N'].Value

    # [a] => [a], base
    if ([string]::IsNullOrWhiteSpace($num)) {
        return @($AltKey, "")
    }

    # [a2] => [a2], [a], base
    return @($AltKey, "[$letter]", "")
}

# --- FUNCTION: Get-NonBangTagsKey ---
# PURPOSE:
# - Produce a tag key that ignores [!] so sets can match across files that differ only by [!].
# NOTES:
# - Used for “relaxed” compatibility matching during selection.
function Get-NonBangTagsKey {
    param([string]$BaseTagsKey)
    if ([string]::IsNullOrWhiteSpace($BaseTagsKey)) { return "" }
    return ($BaseTagsKey -replace '\[\!\]', '')
}

# --- FUNCTION: Normalize-Alt ---
# PURPOSE:
# - Normalize alt tags so blank/whitespace becomes $null for stable comparisons.
# NOTES:
# - Ensures comparisons don’t treat "" and $null as different alts.
function Normalize-Alt {
    param([AllowNull()][AllowEmptyString()][string]$Alt)
    if ([string]::IsNullOrWhiteSpace($Alt)) { return $null }
    return $Alt
}

# --- FUNCTION: Move-FileSafe ---
# PURPOSE:
# - Safe file move helper with a long-path fallback (kept for compatibility).
# NOTES:
# - Not used by the NON-M3U flow (which now edits gamelist.xml instead of moving files).
function Move-FileSafe {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Destination
    )

    # Guard: nothing to do if the same path
    if ($Source -ieq $Destination) { return }

    # Ensure destination directory exists
    $destDir = Split-Path -Parent $Destination
    if (-not (Test-Path -LiteralPath $destDir)) {
        New-Item -ItemType Directory -Path $destDir | Out-Null
    }

    # Attempt normal move first
    try {
        Move-Item -LiteralPath $Source -Destination $Destination -Force -ErrorAction Stop
        return
    } catch {
        # Fall back to long-path move
    }

    # Long path prefix for Windows
    $srcLp = "\\?\$Source"
    $dstLp = "\\?\$Destination"

    # If destination exists, remove before moving
    if ([System.IO.File]::Exists($dstLp)) {
        [System.IO.File]::Delete($dstLp)
    }
    [System.IO.File]::Move($srcLp, $dstLp)
}

# --- FUNCTION: Get-PlatformRootName ---
# PURPOSE:
# - Determine the platform folder name (e.g., "apple2") based on script location and file directory.
# NOTES:
# - If script is run from roms root, platform is the first child folder under roms.
# - If script is run inside a platform folder, that folder name is the platform.
function Get-PlatformRootName {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$ScriptDir
    )

    # Resolve absolute paths for robust comparisons
    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $dirFull    = (Resolve-Path -LiteralPath $Directory).Path.TrimEnd('\')

    # Detect if script is placed at the ROMS root
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # If running from ROMS root, platform is first child folder beneath ROMS
    if ($scriptIsRomsRoot) {
        if (-not $dirFull.StartsWith($scriptFull, [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
        $rel = $dirFull.Substring($scriptFull.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($rel)) { return $null }
        $parts = $rel -split '\\'
        if ($parts.Count -ge 1) { return $parts[0].ToLowerInvariant() }
        return $null
    }

    # If running inside a platform folder, the script folder itself is the platform
    return $scriptLeaf.ToLowerInvariant()
}

# --- FUNCTION: Get-PlatformRootPath ---
# PURPOSE:
# - Resolve the platform root path (roms\<platform>) depending on where the script is running.
# NOTES:
# - Mirrors Get-PlatformRootName assumptions about “roms root” vs “platform folder”.
function Get-PlatformRootPath {
    param(
        [Parameter(Mandatory=$true)][string]$ScriptDir,
        [Parameter(Mandatory=$true)][string]$PlatformRootName
    )

    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # If script is at ROMS root, platform path is ROMS\<platform>
    if ($scriptIsRomsRoot) {
        return (Join-Path $scriptFull $PlatformRootName)
    }

    # If script is inside platform folder, platform root is script folder itself
    return $scriptFull
}

# --- FUNCTION: Get-PlatformCountLabel ---
# PURPOSE:
# - Build a human-friendly label for summary counts (platform + subfolder where applicable).
# NOTES:
# - Uses uppercase for readability in output.
function Get-PlatformCountLabel {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$ScriptDir
    )

    # Resolve absolute paths for robust comparisons
    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $dirFull    = (Resolve-Path -LiteralPath $Directory).Path.TrimEnd('\')

    # If the directory isn't under the script folder, just label with leaf folder
    if (-not $dirFull.StartsWith($scriptFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return (Split-Path -Leaf $dirFull).ToUpperInvariant()
    }

    # Compute relative path segments beneath script folder
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $rel = $dirFull.Substring($scriptFull.Length).TrimStart('\')
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($rel)) { $parts = $rel -split '\\' }

    # Determine whether we're running at ROMS root
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # ROMS root mode: first segment is platform; rest is optional subfolders
    if ($scriptIsRomsRoot) {
        if ($parts.Count -eq 0) { return $scriptLeaf.ToUpperInvariant() }
        $platform = $parts[0].ToUpperInvariant()
        $subParts = if ($parts.Count -gt 1) { $parts[1..($parts.Count-1)] } else { @() }
        if ($subParts.Count -gt 0) { return ($platform + "\" + ($subParts -join "\")) }
        return $platform
    }

    # Platform folder mode: platform is script leaf; include any subfolders beneath it
    $platform = $scriptLeaf.ToUpperInvariant()
    if ($parts.Count -gt 0) { return ($platform + "\" + ($parts -join "\")) }
    return $platform
}

# --- FUNCTION: Get-RelativeGamelistPath ---
# PURPOSE:
# - Convert a ROM full path into the gamelist.xml <path> format: ./subdir/file.ext with / separators.
# NOTES:
# - Returns $null if the ROM path is not under the platform root.
function Get-RelativeGamelistPath {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformRootPath,
        [Parameter(Mandatory=$true)][string]$FileFullPath
    )

    # Resolve both paths; if resolution fails, return $null
    try {
        $rootFull = (Resolve-Path -LiteralPath $PlatformRootPath).Path.TrimEnd('\')
        $fileFull = (Resolve-Path -LiteralPath $FileFullPath).Path
    } catch {
        return $null
    }

    # Ensure the file lives under the platform root
    if (-not $fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    # Convert to relative path and normalize separators to forward slashes
    $rel = $fileFull.Substring($rootFull.Length).TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($rel)) { return $null }

    $rel = $rel -replace '\\', '/'
    return ("./" + $rel)
}

# --- FUNCTION: Get-UniqueGamelistBackupPath ---
# PURPOSE:
# - Select a non-colliding backup filename for gamelist.xml (gamelist.backup, gamelist.backup (1), ...).
# NOTES:
# - Ensures backups are never overwritten across runs.
function Get-UniqueGamelistBackupPath {
    param([Parameter(Mandatory=$true)][string]$GamelistPath)

    # Establish base backup filename in same directory as gamelist.xml
    $dir = Split-Path -Parent $GamelistPath
    $base = Join-Path $dir "gamelist.backup"
    if (-not (Test-Path -LiteralPath $base)) { return $base }

    # If base exists, search for the next available (N) suffix
    $i = 1
    while ($true) {
        $p = Join-Path $dir ("gamelist.backup ({0})" -f $i)
        if (-not (Test-Path -LiteralPath $p)) { return $p }
        $i++
    }
}

# --- FUNCTION: Ensure-GamelistLoaded ---
# PURPOSE:
# - Load and cache gamelist.xml state (lines + changed flag) per platform.
# NOTES:
# - Caches state in $gamelistStateByPlatform to avoid re-reading the file repeatedly.
# - If file can’t be read, Lines stays $null and Exists stays $true.
function Ensure-GamelistLoaded {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformRootLower,
        [Parameter(Mandatory=$true)][string]$PlatformRootPath
    )

    # Return cached state if already loaded
    if ($gamelistStateByPlatform.ContainsKey($PlatformRootLower)) {
        return $gamelistStateByPlatform[$PlatformRootLower]
    }

    # Establish gamelist.xml path for this platform
    $gamelistPath = Join-Path $PlatformRootPath "gamelist.xml"

    # Initialize state object
    $state = [PSCustomObject]@{
        RootPath      = $PlatformRootPath
        GamelistPath  = $gamelistPath
        Lines         = $null
        Changed       = $false
        Exists        = (Test-Path -LiteralPath $gamelistPath)
    }

    # Attempt to load lines if file exists
    if ($state.Exists) {
        try {
            $state.Lines = [System.IO.File]::ReadAllLines($gamelistPath)
        } catch {
            $state.Lines = $null
        }
    }

    # Cache and return
    $gamelistStateByPlatform[$PlatformRootLower] = $state
    return $state
}

# --- FUNCTION: Save-GamelistIfChanged ---
# PURPOSE:
# - Write gamelist.xml back to disk when it was modified, and create one backup per run per gamelist.
# NOTES:
# - Honors $dryRun (no disk writes).
# - Tracks backups per run via $gamelistBackupDone so only one backup is created per gamelist path.
function Save-GamelistIfChanged {
    param([Parameter(Mandatory=$true)]$State)

    # Guard: only save if file exists and was changed and has loaded lines
    if (-not $State.Exists -or -not $State.Changed -or $null -eq $State.Lines) { return $false }

    # Dry-run: report only (no writes)
    if ($dryRun) { return $false }

    # Create a single backup per run per gamelist file
    if (-not $gamelistBackupDone.ContainsKey($State.GamelistPath)) {
        $backupPath = Get-UniqueGamelistBackupPath -GamelistPath $State.GamelistPath
        Copy-Item -LiteralPath $State.GamelistPath -Destination $backupPath -Force
        $gamelistBackupDone[$State.GamelistPath] = $true
    }

    # Write back as UTF-8 without BOM, preserving line breaks
    $text = ($State.Lines -join [Environment]::NewLine)
    [System.IO.File]::WriteAllText($State.GamelistPath, $text, [System.Text.UTF8Encoding]::new($false))
    return $true
}

# --- FUNCTION: Hide-GameEntriesInGamelist ---
# PURPOSE:
# - For NON-M3U platforms, locate each target <path> in gamelist.xml and ensure <hidden>true</hidden>.
# NOTES:
# - Modifies $State.Lines and sets $State.Changed when edits occur (unless $dryRun).
# - Adds reporting entries into the ArrayList buckets: newly hidden, already hidden, or missing.
# - Marks handled files as “used” in $UsedFiles to prevent false “skipped” reports.
function Hide-GameEntriesInGamelist {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)]$Targets,
        [Parameter(Mandatory=$true)][hashtable]$UsedFiles
    )

    # Establish result counters for this hide operation batch
    $result = [PSCustomObject]@{
        DidWork            = $false
        NewlyHiddenCount   = 0
        AlreadyHiddenCount = 0
        MissingCount       = 0
    }

    # If gamelist doesn't exist (or couldn't load), all targets are missing
    if (-not $State.Exists -or $null -eq $State.Lines) {

        foreach ($t in $Targets) {
            # Track missing lookups for later skip-reason logic
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true

            # Add to reporting bucket (ArrayList.Add avoids op_Addition errors)
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })

            $result.MissingCount++
        }

        return $result
    }

    # Establish a working copy reference to lines
    $lines = $State.Lines

    # Process each file target that should be hidden
    foreach ($t in $Targets) {

        # Establish the expected <path> value in gamelist.xml (./rel/path with / separators)
        $rel = $t.RelPath
        if ([string]::IsNullOrWhiteSpace($rel)) {

            # If we can't compute the rel path, treat as missing in gamelist.xml
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
            continue
        }

        # Track whether this target was found and whether we applied/confirmed hidden
        $found = $false
        $handled = $false

        # Walk lines looking for a <path> that matches this rel path
        for ($i = 0; $i -lt $lines.Count; $i++) {

            $line = $lines[$i]
            if ($null -eq $line) { continue }

            # Match: <path>./whatever</path> (capture indentation for insertion)
            $m = [regex]::Match($line, '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
            if (-not $m.Success) { continue }

            $val = $m.Groups['V'].Value
            if ($val -ne $rel) { continue }

            # Found matching <path> entry
            $found = $true
            $indent = $m.Groups['I'].Value

            # Scan forward within this <game> block for an existing <hidden> node
            $j = $i + 1
            $hiddenLineIndex = $null
            $hiddenValue = $null

            while ($j -lt $lines.Count) {
                $tline = $lines[$j]

                # Stop if next game path begins or game ends
                if ($tline -match '^\s*<path>\s*') { break }
                if ($tline -match '^\s*</game>\s*$') { break }

                # Match: <hidden>...</hidden>
                $hm = [regex]::Match($tline, '^\s*<hidden>\s*(?<H>.*?)\s*</hidden>\s*$')
                if ($hm.Success) {
                    $hiddenLineIndex = $j
                    $hiddenValue = $hm.Groups['H'].Value
                    break
                }
                $j++
            }

            # Case 1: <hidden> exists already
            if ($null -ne $hiddenLineIndex) {

                # Subcase: already hidden => count/report, mark used
                if ($hiddenValue -match '^(?i)true$') {
                    $UsedFiles[$t.FullPath] = $true
                    [void]$noM3UAlreadyHidden.Add([PSCustomObject]@{
                        FullPath = $t.FullPath
                        Reason   = "Already hidden in gamelist.xml"
                    })
                    $result.AlreadyHiddenCount++
                    $handled = $true
                }
                # Subcase: hidden exists but not true => flip to true, mark changed (unless dryrun)
                else {
                    if (-not $dryRun) {
                        $lines[$hiddenLineIndex] = ($lines[$hiddenLineIndex] -replace '(?i)<hidden>\s*.*?\s*</hidden>', '<hidden>true</hidden>')
                    }
                    $State.Changed = $true
                    $result.DidWork = $true
                    $result.NewlyHiddenCount++
                    $UsedFiles[$t.FullPath] = $true
                    [void]$noM3UNewlyHidden.Add([PSCustomObject]@{
                        FullPath = $t.FullPath
                        Reason   = "Hidden in gamelist.xml"
                    })
                    $handled = $true
                }
            }
            # Case 2: <hidden> does not exist => insert a new <hidden>true</hidden> line after <path>
            else {

                if (-not $dryRun) {

                    # Build the inserted line at the same indentation level as <path>
                    $insertLine = ($indent + "<hidden>true</hidden>")

                    # Splice-in: lines[0..i] + insertLine + lines[i+1..end]
                    $before = @()
                    if ($i -ge 0) { $before = $lines[0..$i] }
                    $after = @()
                    if (($i + 1) -le ($lines.Count - 1)) { $after = $lines[($i + 1)..($lines.Count - 1)] }

                    $lines = @($before + $insertLine + $after)
                }

                # Mark state changed and report this as a newly hidden entry
                $State.Changed = $true
                $result.DidWork = $true
                $result.NewlyHiddenCount++
                $UsedFiles[$t.FullPath] = $true
                [void]$noM3UNewlyHidden.Add([PSCustomObject]@{
                    FullPath = $t.FullPath
                    Reason   = "Hidden in gamelist.xml"
                })

                $handled = $true

                # Advance index so we don't re-process inserted line in this loop
                $i++
            }

            # Once handled (hidden confirmed or applied), stop scanning for this target
            if ($handled) { break }
        }

        # If we never found a matching <path> entry, mark as missing
        if (-not $found) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
        }
    }

    # Persist the potentially re-built array back into state (important when we spliced lines)
    $State.Lines = $lines
    return $result
}

# --- FUNCTION: Parse-GameFile ---
# PURPOSE:
# - Parse a filename into multi-disk metadata (disk/side/total/tags) used for grouping and selection.
# NOTES:
# - Returns $null for files that do not look like multi-disk candidates.
# - Supports Disk/Disc + numeric/letter/roman tokens, optional “of N”, and Side markers.
# - Side-only patterns are treated as Disk 1 with Side A/B/etc.
function Parse-GameFile {
    param(
        [Parameter(Mandatory=$true)][string]$FileName,
        [Parameter(Mandatory=$true)][string]$Directory
    )

    # Strip extension for token parsing
    $nameNoExt = $FileName -replace '\.[^\.]+$', ''

    # Disk/Disc pattern with optional: "of N" and "Side X"
    $diskPattern = '(?i)^(?<Prefix>.*?)(?<Sep>[\s._\-\(]|^)(?<Type>disk|disc)(?!s)[\s_]*(?<Disk>\d+|[A-Za-z]|(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX))(?:(?=\s+of\s+\d+)|(?=\s+Side\s+[A-Za-z])|(?=[\s\)\]\._-]|$))(?:\s+of\s+(?<Total>\d+))?(?:\s+Side\s+(?<Side>[A-Za-z]))?(?<After>.*)$'

    # Side-only pattern (Side A/B/etc.), treated as Disk 1 with differing sides
    $sideOnlyPattern = '(?i)^(?<Prefix>.*?)(?<Sep>[\s._\-\(]|^)Side\s+(?<SideOnly>[A-Za-z])(?=[\s\)\]\._-]|$)(?<After>.*)$'

    # Attempt disk match first
    $diskMatch = [regex]::Match($nameNoExt, $diskPattern)
    $sideOnlyMatch = $null

    $hasDisk = $diskMatch.Success
    if (-not $hasDisk) {
        $sideOnlyMatch = [regex]::Match($nameNoExt, $sideOnlyPattern)
    }

    # If neither matches, file isn't a multi-disk candidate
    if (-not $hasDisk -and (-not $sideOnlyMatch.Success)) {
        return $null
    }

    # Initialize parsed fields
    $prefixRaw  = ""
    $diskToken  = $null
    $totalToken = $null
    $sideToken  = $null
    $after      = ""

    # Populate fields based on which pattern matched
    if ($hasDisk) {
        $prefixRaw = $diskMatch.Groups['Prefix'].Value
        $diskToken = $diskMatch.Groups['Disk'].Value
        if ($diskMatch.Groups['Total'].Success) { $totalToken = $diskMatch.Groups['Total'].Value }
        if ($diskMatch.Groups['Side'].Success)  { $sideToken  = $diskMatch.Groups['Side'].Value }
        $after = $diskMatch.Groups['After'].Value
    } else {
        $prefixRaw  = $sideOnlyMatch.Groups['Prefix'].Value
        $diskToken  = "1"
        $totalToken = $null
        $sideToken  = $sideOnlyMatch.Groups['SideOnly'].Value
        $after      = $sideOnlyMatch.Groups['After'].Value
    }

    # Normalize base prefix (for grouping + naming)
    $basePrefix = Clean-BasePrefix $prefixRaw

    # Normalize the remainder (strip leading ')' and whitespace artifacts)
    $afterNorm = $after -replace '^[\)\s]+', ''

    # Extract name hint (single leading parenthetical right after the designator), if present
    $nameHint = ""
    $beforeBracket = $afterNorm
    $bracketIdx = $beforeBracket.IndexOf('[')
    if ($bracketIdx -ge 0) { $beforeBracket = $beforeBracket.Substring(0, $bracketIdx) }
    $mHint = [regex]::Match($beforeBracket, '^\s*(\([^\)]+\))')
    if ($mHint.Success) { $nameHint = $mHint.Groups[1].Value }

    # Gather bracket tags (excluding disk-noise tags)
    $bracketTags = @()
    foreach ($m in [regex]::Matches($afterNorm, '\[[^\]]+\]')) {
        $tag = $m.Value
        if (-not (Is-DiskNoiseTag $tag)) { $bracketTags += $tag }
    }

    # Split tags into alt tag (single) and base tags (remaining)
    $altTag  = ""
    $baseTags = @()
    foreach ($t in $bracketTags) {
        if (Is-AltTag $t) { $altTag = $t }
        else { $baseTags += $t }
    }

    # Create stable keys for grouping
    $baseTagsKey = ($baseTags -join "")

    # Convert tokens to sortable values
    $diskSort = Convert-DiskToSort $diskToken
    $sideSort = Convert-SideToSort $sideToken

    # Parse total disks if present
    $totalDisks = $null
    if ($totalToken -match '^\d+$') { $totalDisks = [int]$totalToken }

    # Return parsed metadata object
    return [PSCustomObject]@{
        FileName        = $FileName
        Directory       = $Directory
        BasePrefix      = $basePrefix
        BaseTagsKey     = $baseTagsKey
        BaseTags        = $baseTags
        BaseTagsKeyNB   = (Get-NonBangTagsKey $baseTagsKey)
        AltTag          = $altTag
        DiskSort        = $diskSort
        SideSort        = $sideSort
        TotalDisks      = $totalDisks
        NameHint        = $nameHint
        TitleKey        = ($Directory + "`0" + $basePrefix)
    }
}

# --- FUNCTION: Select-DiskEntries ---
# PURPOSE:
# - Select candidate entries for a specific disk number within a set (optionally constrained by alt and total).
# NOTES:
# - Sorts matches by SideSort so Side A comes before Side B for the same disk.
function Select-DiskEntries {
    param(
        [Parameter(Mandatory=$true)]$Files,
        [Parameter(Mandatory=$true)][int]$DiskNumber,
        [Parameter(Mandatory=$false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$AltTag = "",
        [Parameter(Mandatory=$false)]$RootTotal
    )

    # Normalize requested alt so "" and $null compare consistently
    $wantAlt = Normalize-Alt $AltTag

    # Filter by disk number, alt, and (optionally) total disk count consistency
    $picked = @(
        $Files | Where-Object {
            $_.DiskSort -eq $DiskNumber -and
            (Normalize-Alt $_.AltTag) -eq $wantAlt -and
            ( (-not $RootTotal) -or ($_.TotalDisks -eq $RootTotal) -or ($_.TotalDisks -eq $null) )
        } | Sort-Object SideSort
    )

    return ,$picked
}

# ==================================================================================================
# PHASE 1: FILE ENUMERATION / PARSING
# ==================================================================================================

# Establish collection for parsed candidates
$parsed = @()

# Announce scan phase (Get-ChildItem -Recurse is commonly the longest wait)
Write-Phase "Collecting ROM file data (scanning folders, which might take a while)..."

# Enumerate files under scriptDir (up to 2 subdirectory levels)
Get-ChildItem -Path $scriptDir -File -Recurse -Depth 2 | ForEach-Object {

    # Skip M3U input files (avoid treating playlists as ROM inputs)
    if ($_.Extension -ieq ".m3u") { return }

    # Skip known media/manual folders to reduce false positives
    if ($skipFolders -contains $_.Directory.Name.ToLowerInvariant()) { return }

    # If user wants to skip non-M3U platforms entirely, filter them out during enumeration
    if ($noM3UPlatformMode -ieq "skip") {

        # Determine platform folder for this file
        $plat = Get-PlatformRootName -Directory $_.DirectoryName -ScriptDir $scriptDir
        if ($null -ne $plat) {

            # Normalize configured non-M3U platform list to lowercase for consistent comparisons
            $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
            if ($noM3USetLower -contains $plat.ToLowerInvariant()) { return }
        }
    }

    # Attempt to parse this file as a multi-disk candidate
    $p = Parse-GameFile -FileName $_.Name -Directory $_.DirectoryName
    if ($null -ne $p) { $parsed += $p }
}

# ==================================================================================================
# PHASE 2: INDEXING / GROUPING
# ==================================================================================================

# Announce indexing phase
Write-Phase "Indexing parsed candidates (grouping titles, tags, variants, etc.)..."

# Group strictly by directory + base title prefix + full base tags key (excluding alt tags)
$groupsStrict = $parsed | Group-Object Directory, BasePrefix, BaseTagsKey

# Build a title index: TitleKey -> list of parsed entries (used for relaxed matching)
$titleIndex = @{}
foreach ($p in $parsed) {
    if (-not $titleIndex.ContainsKey($p.TitleKey)) { $titleIndex[$p.TitleKey] = @() }
    $titleIndex[$p.TitleKey] += $p
}

# Track "bang" ([!]) presence per title + non-bang tags, and per title + non-bang tags + alt
$bangByTitleNB = @{}
$bangAltByTitleNB = @{}

foreach ($p in $parsed) {

    # Only record if [!] exists in this entry's base tags key
    if ($p.BaseTagsKey -match '\[\!\]') {

        # TitleKey + NonBangKey indicates at least one [!] variant exists for that title+tag family
        $k = $p.TitleKey + "`0" + $p.BaseTagsKeyNB
        $bangByTitleNB[$k] = $true

        # Also track alt-specific [!] existence (used for suppression rules)
        if (-not [string]::IsNullOrWhiteSpace($p.AltTag)) {
            $k2 = $p.TitleKey + "`0" + $p.BaseTagsKeyNB + "`0" + $p.AltTag
            $bangAltByTitleNB[$k2] = $true
        }
    }
}

# Track playlist paths claimed in this run (prevents collisions)
$occupiedPlaylistPaths = @{}

# Track which playlists were actually written in this run (M3U-only)
$m3uWrittenPlaylistPaths = @{}

# Track duplicate-content signatures to suppress identical playlists in the same run
$playlistSignatures = @{}

# Track suppression buckets for reporting
$suppressedDuplicatePlaylists   = @{} # playlistPath -> collidedWithPath
$suppressedPreExistingPlaylists = @{} # playlistPath -> $true (content identical)
$overwrittenExistingPlaylists   = @{} # playlistPath -> $true (content differed)

# Track which disk files were used by either written/suppressed playlists or successful gamelist hides
$usedFiles = @{}

# ==================================================================================================
# PHASE 3: PROCESS MULTI-DISK GROUPS (M3U PLAYLISTS + NON-M3U GAMELIST HIDING)
# ==================================================================================================

# Announce processing phase
Write-Phase "Processing multi-disk candidates (playlists / gamelist updates)..."

foreach ($group in $groupsStrict) {

    # Establish group context values (shared by all members)
    $groupFiles = $group.Group
    $directory  = $groupFiles[0].Directory
    $basePrefix = $groupFiles[0].BasePrefix
    $titleKey   = $groupFiles[0].TitleKey

    # Establish relaxed candidate list for this title (same title prefix across tags)
    $titleFiles = if ($titleIndex.ContainsKey($titleKey)) { $titleIndex[$titleKey] } else { $groupFiles }

    # Establish non-bang tags key for relaxed matching
    $strictNBKey = $groupFiles[0].BaseTagsKeyNB

    # Filter to title-compatible entries that match non-bang tags key
    $titleCompatible = @(
        $titleFiles | Where-Object { $_.BaseTagsKeyNB -eq $strictNBKey }
    )

    # Establish all alt keys present in this strict group ("" included for base)
    $altKeys = @(
        ($groupFiles | Select-Object -ExpandProperty AltTag | ForEach-Object { if ($_ -eq $null) { "" } else { $_ } } | Sort-Object -Unique)
    )

    foreach ($altKey in $altKeys) {

        # Disk 1 roots help determine "total disk count" variants
        $disk1Roots = @($titleCompatible | Where-Object { $_.DiskSort -eq 1 })

        # Gather explicit totals (Disk 1 "of N") values; fallback to $null if none
        $totals = @($disk1Roots | Where-Object { $_.TotalDisks -ne $null } | Select-Object -ExpandProperty TotalDisks | Sort-Object -Unique)
        if ($totals.Count -eq 0) { $totals = @($null) }

        foreach ($rootTotal in $totals) {

            # Determine which disk numbers we are targeting for this set
            $diskTargets = if ($rootTotal) { 1..$rootTotal }
                           else { @($titleCompatible | Select-Object -ExpandProperty DiskSort | Sort-Object -Unique) }

            # Accumulate selected files for the playlist (or non-M3U set)
            $playlistFiles = @()

            foreach ($d in $diskTargets) {

                # Establish selection list for this disk number
                $picked = @()

                # Pass 1: strict match within the strict group (same dir + base tags + alt)
                $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $altKey -RootTotal $rootTotal

                # Pass 1.5: if base playlist and strict missing, allow a single unambiguous alt disk
                if ($picked.Count -eq 0 -and [string]::IsNullOrWhiteSpace($altKey)) {

                    $candAlt = @(
                        $groupFiles | Where-Object {
                            $_.DiskSort -eq $d -and
                            (-not [string]::IsNullOrWhiteSpace($_.AltTag)) -and
                            ( (-not $rootTotal) -or ($_.TotalDisks -eq $rootTotal) -or ($_.TotalDisks -eq $null) )
                        } | Sort-Object SideSort
                    )

                    if ($candAlt.Count -eq 1) { $picked = $candAlt }
                }

                # Pass 2: relaxed match within the title-compatible family (same title + non-bang tags)
                if ($picked.Count -eq 0) {

                    $wantAlt = Normalize-Alt $altKey
                    $cand = @(
                        $titleCompatible | Where-Object {
                            $_.DiskSort -eq $d -and
                            (Normalize-Alt $_.AltTag) -eq $wantAlt -and
                            ( (-not $rootTotal) -or ($_.TotalDisks -eq $rootTotal) -or ($_.TotalDisks -eq $null) )
                        } | Sort-Object SideSort
                    )

                    if ($cand.Count -gt 0) { $picked = $cand }
                }

                # Pass 3: strict alt fallback chain within the strict group
                if ($picked.Count -eq 0) {
                    $altChain = Get-AltFallbackChain $altKey
                    foreach ($tryAlt in $altChain) {
                        $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $tryAlt -RootTotal $rootTotal
                        if ($picked.Count -gt 0) { break }
                    }
                }

                # Pass 4: relaxed alt fallback chain within the title-compatible family
                if ($picked.Count -eq 0) {
                    $altChain = Get-AltFallbackChain $altKey
                    foreach ($tryAlt in $altChain) {

                        $wantAlt2 = Normalize-Alt $tryAlt
                        $cand = @(
                            $titleCompatible | Where-Object {
                                $_.DiskSort -eq $d -and
                                (Normalize-Alt $_.AltTag) -eq $wantAlt2 -and
                                ( (-not $rootTotal) -or ($_.TotalDisks -eq $rootTotal) -or ($_.TotalDisks -eq $null) )
                            } | Sort-Object SideSort
                        )

                        if ($cand.Count -gt 0) { $picked = $cand; break }
                    }
                }

                # If anything was picked for this disk number, add to set
                if ($picked.Count -gt 0) { $playlistFiles += $picked }
            }

            # Guard: require at least 2 entries to be considered a multi-disk set
            if (@($playlistFiles).Count -lt 2) { continue }

            # Determine whether a single shared name hint exists among all selected entries
            $uniqueHints = @($playlistFiles | Select-Object -ExpandProperty NameHint | Sort-Object -Unique)
            $useHint = ""
            if ($uniqueHints.Count -eq 1 -and (-not [string]::IsNullOrWhiteSpace($uniqueHints[0]))) { $useHint = $uniqueHints[0] }

            # Build base playlist name: base title prefix + (optional) shared hint
            $playlistBase = $basePrefix
            if ($useHint) { $playlistBase += $useHint }

            # Add only those base tags that are common across all selected entries
            $commonBaseTags = @()
            if (@($playlistFiles).Count -gt 0) {
                $firstTags = @($playlistFiles[0].BaseTags)
                foreach ($t in $firstTags) {
                    $inAll = $true
                    foreach ($pf in $playlistFiles) {
                        if (-not ($pf.BaseTags -contains $t)) { $inAll = $false; break }
                    }
                    if ($inAll -and (-not ($commonBaseTags -contains $t))) { $commonBaseTags += $t }
                }
            }
            if ($commonBaseTags.Count -gt 0) { $playlistBase += ($commonBaseTags -join "") }

            # If all entries share the same alt, append it
            $altsInPlaylist = @($playlistFiles | ForEach-Object { Normalize-Alt $_.AltTag } | Sort-Object -Unique)
            if ($altsInPlaylist.Count -eq 1 -and $altsInPlaylist[0]) {
                $playlistBase += $altsInPlaylist[0]
            }

            # Final cleanup of the playlist base name
            $playlistBase = $playlistBase -replace '\s{2,}', ' '
            $playlistBase = $playlistBase -replace '[\s._-]*[\(]*$', ''
            $playlistBase = $playlistBase -replace '\(\s*\)', ''
            $playlistBase = $playlistBase.Trim()
            if ([string]::IsNullOrWhiteSpace($playlistBase)) { continue }

            # Determine the platform root for this directory
            $platformRoot = Get-PlatformRootName -Directory $directory -ScriptDir $scriptDir

            # Determine whether this is a NON-M3U platform
            $isNoM3U = $false
            if ($null -ne $platformRoot) {
                $rootLower = $platformRoot.ToLowerInvariant()
                $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
                $isNoM3U = ($noM3USetLower -contains $rootLower)
            }

            # If NON-M3U and mode is skip, do nothing
            if ($isNoM3U -and $noM3UPlatformMode -ieq "skip") { continue }

            # Sort selected files by disk number then side
            $sorted = $playlistFiles | Sort-Object DiskSort, SideSort

            # Determine whether the set is incomplete (only enforce when rootTotal is known)
            $setIsIncomplete = $false
            if ($rootTotal) {
                $presentDisks = @($playlistFiles | Select-Object -ExpandProperty DiskSort | Sort-Object -Unique)
                $missingDisksLocal = @()
                foreach ($ed in (1..$rootTotal)) {
                    if (-not ($presentDisks -contains $ed)) { $missingDisksLocal += $ed }
                }
                if ($missingDisksLocal.Count -gt 0) { $setIsIncomplete = $true }
            }

            # Apply [!] preference suppression rule:
            # If this strict group does not have [!], but there exists a [!] variant for the same title+nonbang+alt, suppress
            $thisHasBang = ($groupFiles[0].BaseTagsKey -match '\[\!\]')
            $wantAltNorm = Normalize-Alt $altKey
            if (-not $thisHasBang -and $wantAltNorm) {
                $kBang    = $titleKey + "`0" + $strictNBKey
                $kBangAlt = $titleKey + "`0" + $strictNBKey + "`0" + $wantAltNorm
                if ($bangByTitleNB.ContainsKey($kBang) -and $bangAltByTitleNB.ContainsKey($kBangAlt)) {
                    continue
                }
            }

            # Build M3U content (filenames only), then sanitize:
            # - Trim end spaces
            # - Remove blank lines
            # - No trailing newline at EOF (WriteAllText with cleanText as-is)
            $newLines = @($sorted | ForEach-Object { $_.FileName })
            $cleanLines = @(
                $newLines |
                    ForEach-Object { $_.TrimEnd() } |
                    Where-Object { $_ -ne "" }
            )

            $cleanText = ($cleanLines -join "`n")
            $newNorm   = Normalize-M3UText $cleanText

            # Compute playlist path:
            # - NON-M3U: path without extension is used as an internal "claimed" key only
            # - M3U: actual .m3u file path
            $playlistPath = if ($isNoM3U) {
                Join-Path $directory "$playlistBase"
            } else {
                Join-Path $directory "$playlistBase.m3u"
            }

            # Collision avoidance: if this path was already claimed during this run, generate [alt] variants
            if ($occupiedPlaylistPaths.ContainsKey($playlistPath)) {
                $altIndex = 1
                do {
                    $suffix = if ($altIndex -eq 1) { "[alt]" } else { "[alt$altIndex]" }
                    if ($isNoM3U) {
                        $playlistPath = Join-Path $directory "$playlistBase$suffix"
                    } else {
                        $playlistPath = Join-Path $directory "$playlistBase$suffix.m3u"
                    }
                    $altIndex++
                } while ($occupiedPlaylistPaths.ContainsKey($playlistPath))
            }

            # Build signature for duplicate-content suppression during the same run
            $sigParts = @()
            foreach ($sf in $sorted) { $sigParts += (Join-Path $sf.Directory $sf.FileName) }
            $playlistSig = ($sigParts -join "`0")

            # Suppress identical content duplicates in the same run
            if ($playlistSignatures.ContainsKey($playlistSig)) {

                # Record suppression reason
                $suppressedDuplicatePlaylists[$playlistPath] = $playlistSignatures[$playlistSig]
                $occupiedPlaylistPaths[$playlistPath] = $true

                # Mark all these files as used (so they don't appear as skipped)
                foreach ($sf in $sorted) {
                    $full = Join-Path $sf.Directory $sf.FileName
                    $usedFiles[$full] = $true
                }

                continue
            }

            # For M3U output, suppress overwrites if the existing file content is identical (after BOM/newline normalization)
            if (-not $isNoM3U) {

                $existsOnDisk = Test-Path -LiteralPath $playlistPath
                if ($existsOnDisk) {

                    # Read existing playlist text
                    $existingRaw = $null
                    try {
                        $existingRaw = Get-Content -LiteralPath $playlistPath -Raw -ErrorAction Stop
                    } catch {
                        $existingRaw = $null
                    }

                    # Normalize for equality comparison
                    $existingNorm = Normalize-M3UText $existingRaw

                    # Compare line-by-line
                    $sameContent = ($existingNorm.Count -eq $newNorm.Count)
                    if ($sameContent) {
                        for ($i = 0; $i -lt $existingNorm.Count; $i++) {
                            if ($existingNorm[$i] -ne $newNorm[$i]) { $sameContent = $false; break }
                        }
                    }

                    # If identical, suppress write
                    if ($sameContent) {
                        $suppressedPreExistingPlaylists[$playlistPath] = $true
                        $occupiedPlaylistPaths[$playlistPath] = $true
                        $playlistSignatures[$playlistSig] = $playlistPath

                        # Mark these files as used
                        foreach ($sf in $sorted) {
                            $full = Join-Path $sf.Directory $sf.FileName
                            $usedFiles[$full] = $true
                        }

                        continue
                    }

                    # Otherwise: file exists and differs => we'll overwrite it
                    $overwrittenExistingPlaylists[$playlistPath] = $true
                }
            }

            # Record this signature as emitted (written or processed)
            $playlistSignatures[$playlistSig] = $playlistPath
            $occupiedPlaylistPaths[$playlistPath] = $true

            # --------------------------------------------------------------------------------------------------
            # BRANCH: NON-M3U PLATFORMS => GAMELIST HIDING
            # --------------------------------------------------------------------------------------------------
            if ($isNoM3U) {

                # Establish primary entry to keep visible: the first file in sorted set
                $primaryFull = Join-Path $sorted[0].Directory $sorted[0].FileName

                # If the set is incomplete, do not hide anything; keep visible and annotate
                if ($setIsIncomplete) {
                    [void]$noM3UPrimaryEntriesIncomplete.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Primary entry kept visible (disk set incomplete)"
                    })
                    continue
                }

                # Determine entries to hide (Disk 2+ equivalents)
                $toHide = @($sorted | Select-Object -Skip 1)

                # If there's nothing to hide, record primary and move on
                if ($toHide.Count -eq 0) {
                    [void]$noM3UPrimaryEntriesOk.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Primary entry kept visible"
                    })
                    continue
                }

                # Compute platform root + load gamelist state (cached)
                $platformLower = $platformRoot.ToLowerInvariant()
                $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
                $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

                # Build hide targets with relative gamelist paths
                $targets = @()
                foreach ($sf in $toHide) {

                    $fullFile = Join-Path $sf.Directory $sf.FileName
                    $rel = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $fullFile

                    $platformLabel = $platformRoot.ToUpperInvariant()
                    $targets += [PSCustomObject]@{
                        FullPath      = $fullFile
                        RelPath       = $rel
                        PlatformLabel = $platformLabel
                    }
                }

                # Perform hide operations and get result counts
                $hideResult = Hide-GameEntriesInGamelist -State $state -Targets $targets -UsedFiles $usedFiles

                # Accumulate per-platform counts for summary reporting
                $platLabel = $platformRoot.ToUpperInvariant()
                if (-not $gamelistHiddenCounts.ContainsKey($platLabel)) { $gamelistHiddenCounts[$platLabel] = 0 }
                if (-not $gamelistAlreadyHiddenCounts.ContainsKey($platLabel)) { $gamelistAlreadyHiddenCounts[$platLabel] = 0 }
                $gamelistHiddenCounts[$platLabel] += [int]$hideResult.NewlyHiddenCount
                $gamelistAlreadyHiddenCounts[$platLabel] += [int]$hideResult.AlreadyHiddenCount

                # Record the primary entry and annotate if some entries were missing from gamelist.xml
                if ($hideResult.MissingCount -gt 0) {
                    [void]$noM3UPrimaryEntriesOk.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Primary entry kept visible (some set entries missing from gamelist.xml)"
                    })
                } else {
                    [void]$noM3UPrimaryEntriesOk.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Primary entry kept visible"
                    })
                }

                # Count NON-M3U “set handled” as a “playlist created” in the summary counters
                $platformLabel = $platformRoot.ToUpperInvariant()
                if (-not $platformCounts.ContainsKey($platformLabel)) { $platformCounts[$platformLabel] = 0 }
                $platformCounts[$platformLabel]++
                $totalPlaylistsCreated++

                continue
            }

            # --------------------------------------------------------------------------------------------------
            # BRANCH: M3U PLATFORMS => WRITE PLAYLISTS
            # --------------------------------------------------------------------------------------------------

            # If set is incomplete, do not write an M3U playlist
            if ($setIsIncomplete) { continue }

            # Write M3U playlist file unless in dry-run mode
            if (-not $dryRun) {
                [System.IO.File]::WriteAllText($playlistPath, $cleanText, [System.Text.UTF8Encoding]::new($false))
            }

            # Track that this playlist was written (M3U only)
            $m3uWrittenPlaylistPaths[$playlistPath] = $true

            # Mark source files as used
            foreach ($sf in $sorted) {
                $full = Join-Path $sf.Directory $sf.FileName
                $usedFiles[$full] = $true
            }

            # Increment per-platform counts
            $platformLabel = Get-PlatformCountLabel -Directory $directory -ScriptDir $scriptDir
            if (-not $platformCounts.ContainsKey($platformLabel)) { $platformCounts[$platformLabel] = 0 }
            $platformCounts[$platformLabel]++
            $totalPlaylistsCreated++
        }
    }
}

# ==================================================================================================
# PHASE 4: FLUSH GAMELIST CHANGES (NON-M3U PLATFORMS)
# ==================================================================================================

# Announce finalization phase
Write-Phase "Finalizing gamelist.xml updates (if any)..."

# Write any changed gamelist states back to disk (or just mark as changed in dry-run)
foreach ($k in $gamelistStateByPlatform.Keys) {
    $st = $gamelistStateByPlatform[$k]
    if ($null -ne $st) {
        Save-GamelistIfChanged -State $st | Out-Null
    }
}

# ==================================================================================================
# PHASE 5: STRUCTURED REPORTING (SPLIT: M3U PLAYLISTS vs GAMELIST(S) UPDATED)
# ==================================================================================================

# Establish whether we had any M3U-related activity (written or suppressed)
$anyM3UActivity =
    (@($m3uWrittenPlaylistPaths.Keys).Count -gt 0) -or
    (@($suppressedPreExistingPlaylists.Keys).Count -gt 0) -or
    (@($suppressedDuplicatePlaylists.Keys).Count -gt 0)

# Establish whether we had any gamelist-related activity (non-M3U sets identified/hid/missing)
$anyGamelistActivity =
    (@($noM3UPrimaryEntriesOk).Count -gt 0) -or
    (@($noM3UPrimaryEntriesIncomplete).Count -gt 0) -or
    (@($noM3UNewlyHidden).Count -gt 0) -or
    (@($noM3UAlreadyHidden).Count -gt 0) -or
    (@($noM3UMissingGamelistEntries).Count -gt 0) -or
    ($gamelistHiddenCounts.Count -gt 0) -or
    ($gamelistAlreadyHiddenCounts.Count -gt 0)

# If nothing happened at all, show a single clear message and stop
if (-not $anyM3UActivity -and -not $anyGamelistActivity) {
    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green
    Write-Host "No viable multi-disk files were found to create playlists from." -ForegroundColor Yellow
} else {

    # --------------------------------------------------------------------------------------------------
    # SECTION: M3U PLAYLISTS
    # --------------------------------------------------------------------------------------------------
    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green

    # Subsection: created M3U playlists (actually written this run)
    if (@($m3uWrittenPlaylistPaths.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "CREATED" -ForegroundColor Green

        $m3uWrittenPlaylistPaths.Keys | Sort-Object | ForEach-Object {

            $p = $_

            # If we overwrote a differing playlist, highlight it
            if ($overwrittenExistingPlaylists.ContainsKey($p)) {
                Write-Host "$p" -NoNewline
                Write-Host " — Overwrote existing playlist that contained content discrepancy" -ForegroundColor Yellow
            } else {

                # If dry-run, explicitly label the action
                if ($dryRun) {
                    Write-Host "$p" -NoNewline
                    Write-Host " — DRY RUN (would write)" -ForegroundColor Yellow
                } else {
                    Write-Host $p
                }
            }
        }
    }
    # No created M3Us
    else {
        if ($dryRun) {
            Write-Host "No M3U playlists were written (DRY RUN)." -ForegroundColor Yellow
        } else {
            Write-Host "No M3U playlists were written." -ForegroundColor Yellow
        }
    }

    # Subsection: suppressed due to identical pre-existing content
    if (@($suppressedPreExistingPlaylists.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "SUPPRESSED (PRE-EXISTING PLAYLIST CONTAINED IDENTICAL CONTENT)" -ForegroundColor Green
        $suppressedPreExistingPlaylists.Keys | Sort-Object | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
    }

    # Subsection: suppressed due to same-run duplicate collision
    if (@($suppressedDuplicatePlaylists.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "SUPPRESSED (DUPLICATE CONTENT COLLISION DURING THIS RUN)" -ForegroundColor Green
        $suppressedDuplicatePlaylists.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-Host "$($_.Key)" -NoNewline -ForegroundColor Gray
            Write-Host " — Identical content collision with $($_.Value)" -ForegroundColor Yellow
        }
    }

    # --------------------------------------------------------------------------------------------------
    # SECTION: GAMELIST(S) UPDATED (NON-M3U PLATFORMS)
    # --------------------------------------------------------------------------------------------------
    if ($anyGamelistActivity) {
        Write-Host ""

        # Heading reflects dry-run state
        if ($dryRun) {
            Write-Host "GAMELIST(S) UPDATED (DRY RUN — NO FILES MODIFIED)" -ForegroundColor Green
        } else {
            Write-Host "GAMELIST(S) UPDATED" -ForegroundColor Green
        }

        # List which gamelist.xml files were (or would be) modified
        $changedGamelists = @()
        foreach ($k in $gamelistStateByPlatform.Keys) {
            $st = $gamelistStateByPlatform[$k]
            if ($null -ne $st -and $st.Exists -and $st.Changed -and $null -ne $st.GamelistPath) {
                $changedGamelists += $st.GamelistPath
            }
        }

        # Print modified gamelist paths (one per line)
        if ($changedGamelists.Count -gt 0) {
            $changedGamelists | Sort-Object -Unique | ForEach-Object {
                Write-Host "$_" -NoNewline
                if ($dryRun) {
                    Write-Host " — DRY RUN (would modify)" -ForegroundColor Yellow
                } else {
                    Write-Host " — Modified" -ForegroundColor Yellow
                }
            }
        } else {
            # If we had non-M3U activity but no gamelist changes, say so explicitly
            Write-Host "No gamelist.xml files required modification." -ForegroundColor Yellow
        }

        # Subsection: non-M3U sets identified (primary entries kept visible) — OK
        if (@($noM3UPrimaryEntriesOk).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (PRIMARY ENTRIES KEPT VISIBLE — OK)" -ForegroundColor Green
            $noM3UPrimaryEntriesOk | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Subsection: non-M3U sets identified (primary entries kept visible) — SET INCOMPLETE
        if (@($noM3UPrimaryEntriesIncomplete).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (PRIMARY ENTRIES KEPT VISIBLE — SET INCOMPLETE)" -ForegroundColor Green
            $noM3UPrimaryEntriesIncomplete | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Subsection: newly hidden entries
        if (@($noM3UNewlyHidden).Count -gt 0) {
            Write-Host ""
            Write-Host "HIDDEN ENTRIES (NEW)" -ForegroundColor Green
            $noM3UNewlyHidden | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                if ($dryRun) {
                    Write-Host " — DRY RUN (would hide in gamelist.xml)" -ForegroundColor Yellow
                } else {
                    Write-Host " — $($_.Reason)" -ForegroundColor Yellow
                }
            }
        }

        # Subsection: already hidden entries
        if (@($noM3UAlreadyHidden).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES ALREADY HIDDEN (NO CHANGE)" -ForegroundColor Green
            $noM3UAlreadyHidden | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Subsection: missing from gamelist.xml
        if (@($noM3UMissingGamelistEntries).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES MISSING FROM GAMELIST (RUN GAMELIST UPDATE, SCRAPE THE GAME OR MANUALLY ADD IT INTO GAMELIST.XML)" -ForegroundColor Green
            $noM3UMissingGamelistEntries | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }
    }
}

# ==================================================================================================
# PHASE 6: REPORTING — (POSSIBLE) MULTI-DISK FILES SKIPPED
# ==================================================================================================

# Establish skipped-report bucket as ArrayList for stability on large runs
$notInPlaylists = [System.Collections.ArrayList]::new()

# Group by TitleKey to detect orphaned Disk 2+ or incomplete sets
$groupsForNotUsed = $parsed | Group-Object TitleKey
foreach ($g in $groupsForNotUsed) {

    $gFiles = $g.Group

    # --------------------------------------------------------------------------------------------------
    # CASE: Only one file in the title group
    # --------------------------------------------------------------------------------------------------
    if (@($gFiles).Count -lt 2) {

        $f = $gFiles[0]
        $full = Join-Path $f.Directory $f.FileName

        # Determine if this singleton should be reported as an orphan (Disk 2+ or Total >= 2)
        $shouldReportSingleton =
            (($f.TotalDisks -ne $null -and $f.TotalDisks -ge 2) -or
             ($f.DiskSort -ne $null -and $f.DiskSort -ge 2))

        # If it should be reported and was not used, add to skipped list
        if ($shouldReportSingleton -and (-not $usedFiles.ContainsKey($full))) {

            if ($noM3UMissingGamelistEntryByFullPath.ContainsKey($full)) {
                [void]$notInPlaylists.Add([PSCustomObject]@{
                    FullPath = $full
                    Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
                })
            } else {
                [void]$notInPlaylists.Add([PSCustomObject]@{
                    FullPath = $full
                    Reason   = "Incomplete disk set (orphan singleton)"
                })
            }
        }

        continue
    }

    # --------------------------------------------------------------------------------------------------
    # CASE: Multi-file title group — compute expected disks and analyze unused entries
    # --------------------------------------------------------------------------------------------------

    # Determine which disk numbers are present in this title group
    $diskSet = @(
        $gFiles |
            Where-Object { $_.DiskSort -ne $null } |
            Select-Object -ExpandProperty DiskSort |
            Sort-Object -Unique
    )
    if ($diskSet.Count -eq 0) { continue }

    # Determine max explicit total (if any)
    $maxTotal = ($gFiles |
        Where-Object { $_.TotalDisks -ne $null } |
        Select-Object -ExpandProperty TotalDisks |
        Sort-Object -Descending |
        Select-Object -First 1)

    # Build expected disks either from explicit total or from maximum disk seen
    $expectedDisks = @()
    if ($null -ne $maxTotal -and $maxTotal -ne "") { $expectedDisks = 1..([int]$maxTotal) }
    else {
        $maxDisk = ($diskSet | Sort-Object -Descending | Select-Object -First 1)
        if ($null -ne $maxDisk) { $expectedDisks = 1..([int]$maxDisk) } else { $expectedDisks = $diskSet }
    }

    # Build a map of disk -> list of alts present (helps detect alt mismatch issues)
    $altByDisk = @{}
    foreach ($x in $gFiles) {
        if ($null -eq $x.DiskSort) { continue }
        $dsk = [int]$x.DiskSort
        $a = Normalize-Alt $x.AltTag
        if (-not $altByDisk.ContainsKey($dsk)) { $altByDisk[$dsk] = @() }
        if (-not ($altByDisk[$dsk] -contains $a)) { $altByDisk[$dsk] += $a }
    }

    # Determine whether disk totals disagree across files (for reason labeling)
    $totalsDistinct = @(
        $gFiles | Where-Object { $_.TotalDisks -ne $null } |
            Select-Object -ExpandProperty TotalDisks | Sort-Object -Unique
    )
    $minTotal = $null
    $maxTotalLocal = $null
    if ($totalsDistinct.Count -gt 0) {
        $minTotal = ($totalsDistinct | Sort-Object | Select-Object -First 1)
        $maxTotalLocal = ($totalsDistinct | Sort-Object -Descending | Select-Object -First 1)
    }

    # Compute missing disks
    $missingDisks = @()
    foreach ($ed in $expectedDisks) {
        if (-not ($diskSet -contains $ed)) { $missingDisks += $ed }
    }

    # Inspect each file in the title group; report those that were not used
    foreach ($f in $gFiles) {

        $full = Join-Path $f.Directory $f.FileName
        if ($usedFiles.ContainsKey($full)) { continue }

        # If this file is known missing from gamelist.xml for NON-M3U, prefer that reason
        if ($noM3UMissingGamelistEntryByFullPath.ContainsKey($full)) {
            [void]$notInPlaylists.Add([PSCustomObject]@{
                FullPath = $full
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            continue
        }

        # Suppress reporting for non-M3U subfolders (only report root folder issues there)
        $plat = Get-PlatformRootName -Directory $f.Directory -ScriptDir $scriptDir
        if ($null -ne $plat) {

            $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
            if ($noM3USetLower -contains $plat.ToLowerInvariant()) {

                $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $plat
                try {
                    $dirFull  = (Resolve-Path -LiteralPath $f.Directory).Path.TrimEnd('\')
                    $rootFull = (Resolve-Path -LiteralPath $rootPath).Path.TrimEnd('\')
                    if ($dirFull -ne $rootFull) {
                        continue
                    }
                } catch {
                    # If resolution fails, fall through and report rather than silently skipping
                }
            }
        }

        # Default reason (may be refined below)
        $reason = "Unselected during fill"

        # Apply [!] suppression reason if applicable
        $hasBangInThisFile = ($f.BaseTagsKey -match '\[\!\]')
        $altNorm = Normalize-Alt $f.AltTag
        if (-not $hasBangInThisFile -and $altNorm) {

            $kBang = $f.TitleKey + "`0" + $f.BaseTagsKeyNB
            $kBangAlt = $f.TitleKey + "`0" + $f.BaseTagsKeyNB + "`0" + $altNorm

            if ($bangByTitleNB.ContainsKey($kBang) -and $bangAltByTitleNB.ContainsKey($kBangAlt)) {
                $reason = "Suppressed by [!] rule"
            }
        }

        # If disks are missing in this title group, treat as incomplete set
        if ($reason -eq "Unselected during fill" -and $missingDisks.Count -gt 0) {
            $reason = "Missing matching disk"
        }

        # If no disks missing but alt exists, a failure might be due to alt fallback
        if ($reason -eq "Unselected during fill" -and $missingDisks.Count -eq 0 -and $altNorm) {
            $reason = "Alt fallback failed"
        }

        # If alt fallback failed, add more detail when the alt doesn't appear elsewhere across disks
        if ($reason -eq "Alt fallback failed") {
            $aHere = Normalize-Alt $f.AltTag
            if ($aHere) {
                $seenElsewhere = $false
                foreach ($dk in $altByDisk.Keys) {
                    if ([int]$dk -eq [int]$f.DiskSort) { continue }
                    if ($altByDisk[$dk] -contains $aHere) { $seenElsewhere = $true; break }
                }
                if (-not $seenElsewhere) { $reason += " due to alt-tag mismatch across disks" }
            }
        }

        # Normalize "Missing matching disk" to "Incomplete disk set" for final reporting
        if ($reason -eq "Missing matching disk" -and $missingDisks.Count -gt 0) {
            $reason = "Incomplete disk set"
        }

        # If totals disagree and this file appears beyond the minimum total, flag disk total mismatch
        if ($reason -eq "Unselected during fill" -and
            $null -ne $minTotal -and $null -ne $maxTotalLocal -and
            $minTotal -ne $maxTotalLocal -and
            $f.DiskSort -gt [int]$minTotal) {

            $reason += " due to disk total mismatch"
        }

        # Add to skipped-report bucket
        [void]$notInPlaylists.Add([PSCustomObject]@{
            FullPath = $full
            Reason   = $reason
        })
    }
}

# Print skipped report if any entries exist
if (@($notInPlaylists).Count -gt 0) {
    Write-Host ""
    Write-Host "(POSSIBLE) MULTI-DISK FILES SKIPPED" -ForegroundColor Green
    $notInPlaylists | Sort-Object FullPath | ForEach-Object {
        Write-Host "$($_.FullPath)" -NoNewline
        Write-Host " — $($_.Reason)" -ForegroundColor Yellow
    }
}

# ==================================================================================================
# PHASE 7: SUMMARY COUNTS
# ==================================================================================================

# --- REPORT: PLAYLIST CREATION COUNT(S) ---
if ($platformCounts.Count -gt 0 -or $totalPlaylistsCreated -gt 0) {
    Write-Host ""
    Write-Host "PLAYLIST CREATION COUNT(S)" -ForegroundColor Green
    $platformCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Name):" -ForegroundColor Cyan -NoNewline
        Write-Host " $($_.Value)"
    }
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $totalPlaylistsCreated"
}

# --- REPORT: GAMELIST HIDDEN ENTRY COUNT(S) ---
$gamelistHiddenTotal = 0
if ($gamelistHiddenCounts.Count -gt 0) {
    Write-Host ""
    Write-Host "GAMELIST HIDDEN ENTRY COUNT(S)" -ForegroundColor Green
    $gamelistHiddenCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        $gamelistHiddenTotal += [int]$_.Value
        Write-Host "$($_.Name):" -ForegroundColor Cyan -NoNewline
        Write-Host " $($_.Value)"
    }
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $gamelistHiddenTotal"
}

# --- REPORT: GAMELIST ALREADY-HIDDEN COUNT(S) ---
$gamelistAlreadyHiddenTotal = 0
if ($gamelistAlreadyHiddenCounts.Count -gt 0) {
    Write-Host ""
    Write-Host "GAMELIST ALREADY-HIDDEN COUNT(S)" -ForegroundColor Green
    $gamelistAlreadyHiddenCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        $gamelistAlreadyHiddenTotal += [int]$_.Value
        Write-Host "$($_.Name):" -ForegroundColor Cyan -NoNewline
        Write-Host " $($_.Value)"
    }
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $gamelistAlreadyHiddenTotal"
}

# --- REPORT: SUPPRESSED PLAYLIST COUNT(S) ---
$preExistingSuppressedCount = @($suppressedPreExistingPlaylists.Keys).Count
$collisionSuppressedCount   = @($suppressedDuplicatePlaylists.Keys).Count
$totalSuppressedCount       = $preExistingSuppressedCount + $collisionSuppressedCount

if ($totalSuppressedCount -gt 0) {
    Write-Host ""
    Write-Host "SUPPRESSED PLAYLIST COUNT(S)" -ForegroundColor Green
    Write-Host "PRE-EXISTING:" -ForegroundColor Cyan -NoNewline
    Write-Host " $preExistingSuppressedCount"
    Write-Host "COLLISIONS:" -ForegroundColor Cyan -NoNewline
    Write-Host " $collisionSuppressedCount"
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $totalSuppressedCount"
}

# --- REPORT: SKIPPED MULTI-DISK FILE COUNT ---
if (@($notInPlaylists).Count -gt 0) {
    Write-Host ""
    Write-Host "MULTI-DISK FILE SKIP COUNT" -ForegroundColor Green
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $(@($notInPlaylists).Count)"
}

# ==================================================================================================
# PHASE 8: FINAL RUNTIME REPORT
# ==================================================================================================

# Compute elapsed runtime
$elapsed = (Get-Date) - $scriptStart
$totalSeconds = [int][math]::Floor($elapsed.TotalSeconds)

# Format runtime as:
# - X seconds (<60)
# - M:SS (<60m)
# - H:MM:SS (>=60m)
$runtimeText = ""
if ($totalSeconds -lt 60) {
    $runtimeText = "$totalSeconds seconds"
} elseif ($totalSeconds -lt 3600) {
    $m = [int]($totalSeconds / 60)
    $s = $totalSeconds % 60
    $runtimeText = "{0}:{1:D2}" -f $m, $s
} else {
    $h = [int]($totalSeconds / 3600)
    $rem = $totalSeconds % 3600
    $m = [int]($rem / 60)
    $s = $rem % 60
    $runtimeText = "{0}:{1:D2}:{2:D2}" -f $h, $m, $s
}

# Print runtime
Write-Host ""
Write-Host "Runtime:" -ForegroundColor White -NoNewline
Write-Host " $runtimeText"
