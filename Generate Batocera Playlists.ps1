<#
PURPOSE: Create .m3u files for each multi-disk game and insert the list of game filenames into the playlist or update gamelist.xml
VERSION: 1.6
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Requires PowerShell version 7+
- Place this file into the ROMS folder to process all platforms, or in a platform's individual subfolder to process just that one.
- False detections and misses are possible, especially for complex naming structures, but should be rare.
    - Built-in intelligence attempts to determine and annotate why:
        - A playlist might have been suppressed
        - A multi-disk file wasn't incorporated into a playlist
- There are three variables you can update to change script behavior:
    - $dryRun : default is $false; set to $true to preview output without modifying files
    - $nonM3UPlatforms : array of platforms (folder names) that should NOT use M3U; default: 3DO + apple2
    - $noM3UPlatformMode : default "XML"; set to "skip" to ignore NON-M3U platforms entirely

BREAKDOWN
- Enumerates ROM game files starting in the directory the script resides in
    - Scans up to 2 subdirectory levels deep recursively
    - Skips .m3u files during scanning (so it doesn’t treat playlists as input)
    - Skips common media/manual folders (e.g., images, videos, media, manuals, downloaded_*) to reduce false multi-disk detections
    - For platforms that can't use M3U playlist files, the script instead hides Disk 2+ in gamelist.xml (<hidden>true</hidden>)
        - Per run, a backup of gamelist.xml is made first called gamelist.backup, labeled with (1), (2), etc. if a backup already exists.
        - This creates a single entry in Batocera for a multi-disk game (for the first disk in the set) instead of one for each disk
        - Initially, this includes 3DO and Apple II but additional platform folders can be added into the $nonM3UPlatforms array
        - Ensures the canonical <name> exists across Disk 2+ entries
            - If Disk 1 already has a <name> in gamelist.xml, that name is treated as authoritative and is propagated to all discs in the set
        - If you'd rather just skip these platforms, change the $noM3UPlatformMode variable to "skip" instead of "XML"
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
    - Ignores bracket tags that simply mirror the file extension (e.g., [nib] on .nib) so they do not alter grouping or playlist naming.
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
    - For NON-M3U platforms, entries that are unhidden due to reclassification/incompleteness are marked “used”
    - Remaining parsed multi-disk candidates that weren’t used are reported as:
        - (POSSIBLE) MULTI-DISK FILES SKIPPED, with a reason such as:
        - incomplete disk set
        - missing matching disk
        - suppressed by [!] preference rule
        - alt fallback issues
        - disk total mismatch issues
        - no entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)
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
# PHASE 0: SCRIPT STARTUP / CONFIG / STATE
# ==================================================================================================

# [PHASE 0] Establish script working directory and start time
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptStart = Get-Date

# [PHASE 0] Initialize platform playlist count buckets and total counter
$platformCounts = @{}
$totalPlaylistsCreated = 0

# [PHASE 0] User config: dry run toggle
$dryRun = $false # <-- set to $true if you want to see what the output would be without changing files

# [PHASE 0] User config: platforms that should not use M3U playlists
$nonM3UPlatforms = @(
    '3DO'
    'apple2'
)

# [PHASE 0] User config: NON-M3U platform handling mode
$noM3UPlatformMode = "XML"   # <-- set to "skip" to completely ignore those platforms

# [PHASE 0] Attempt to widen console buffer (best-effort)
try {
    $raw = $Host.UI.RawUI
    $size = $raw.BufferSize
    # [PHASE 0] If buffer width is narrow, expand it
    if ($size.Width -lt 300) {
        $raw.BufferSize = New-Object Management.Automation.Host.Size(300, $size.Height)
    }
} catch {
    # [PHASE 0] Ignore console resize failures
}

# [PHASE 0] Declare folder names that should not be scanned
$skipFolders = @(
    'images','videos','media','manuals',
    'downloaded_images','downloaded_videos','downloaded_media','downloaded_manuals'
)

# [PHASE 0] Initialize cached gamelist state per platform
$gamelistStateByPlatform = @{}     # platformLower -> state object (cached lines)

# [PHASE 0] Initialize per-run gamelist backup tracking
$gamelistBackupDone      = @{}     # gamelist path -> $true

# [PHASE 0] Initialize a lookup of "missing from gamelist.xml" by full ROM path
$noM3UMissingGamelistEntryByFullPath = @{}  # full file path -> $true

# [PHASE 0] Initialize NON-M3U reporting buckets (ArrayList for safety)
$noM3UPrimaryEntriesOk         = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UPrimaryEntriesIncomplete = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UNoDisk1Sets              = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

$noM3UNewlyHidden              = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UAlreadyHidden            = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UAlreadyHiddenSet         = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$noM3UNewlyUnhidden            = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UAlreadyVisible           = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UAlreadyVisibleSet        = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

$noM3UMissingGamelistEntries   = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

# [PHASE 0] Initialize per-platform count buckets for gamelist changes
$gamelistHiddenCounts          = @{}  # platform label -> count newly hidden
$gamelistAlreadyHiddenCounts   = @{}  # platform label -> count already hidden (no change)

$gamelistUnhiddenCounts        = @{}  # platform label -> count newly unhidden
$gamelistAlreadyVisibleCounts  = @{}  # platform label -> count already visible (no change)

# [PHASE 0] Track NON-M3U entries that SHOULD be hidden per platform (for Phase 3.5 reconciliation)
$noM3UShouldBeHiddenByPlatform = @{}  # platformLower -> hashtable fullPath -> relPath
$noM3UPlatformsEncountered     = @{}  # platformLower -> $true

# ==================================================================================================
# PHASE 0.5: FUNCTIONS
# ==================================================================================================

# Normalize an M3U payload into a stable line array
function Normalize-M3UText {
    param([AllowNull()][string]$Text)

    # Return empty array for null input
    if ($null -eq $Text) { return @() }

    # Strip BOM if present
    if ($Text.Length -gt 0 -and [int]$Text[0] -eq 0xFEFF) {
        $Text = $Text.Substring(1)
    }

    # Normalize Windows and old-Mac newlines to LF
    $Text = $Text -replace "`r`n", "`n"
    $Text = $Text -replace "`r", "`n"

    # Split into lines
    return ,($Text -split "`n")
}

# Convert disk token (numeric/roman/letter) to sortable integer
function Convert-DiskToSort {
    param([string]$DiskToken)

    # Bail on empty token
    if ([string]::IsNullOrWhiteSpace($DiskToken)) { return $null }

    # Handle numeric tokens
    if ($DiskToken -match '^\d+$') { return [int]$DiskToken }

    # Define roman numeral mapping
    $romanMap = @{
        'I' = 1;  'II' = 2;  'III' = 3;  'IV' = 4;  'V' = 5
        'VI' = 6; 'VII' = 7; 'VIII' = 8; 'IX' = 9;  'X' = 10
        'XI' = 11; 'XII' = 12; 'XIII' = 13; 'XIV' = 14; 'XV' = 15
        'XVI' = 16; 'XVII' = 17; 'XVIII' = 18; 'XIX' = 19; 'XX' = 20
    }

    # Normalize token case
    $upper = $DiskToken.ToUpperInvariant()

    # Handle roman numerals
    if ($romanMap.ContainsKey($upper)) { return $romanMap[$upper] }

    # Handle single letters (A=1, B=2, ...)
    if ($upper -match '^[A-Z]$') {
        return ([int][char]$upper[0]) - 64
    }

    # Unknown token => null
    return $null
}

# Convert side token (A/B/C...) to sortable integer
function Convert-SideToSort {
    param([string]$SideToken)

    # Treat missing side as 0
    if ([string]::IsNullOrWhiteSpace($SideToken)) { return 0 }

    # Convert letter to 1-based index
    $c = $SideToken.ToUpperInvariant()[0]
    return ([int][char]$c) - 64
}

# Identify TOSEC-style alt tags like [a], [a2], [b], [b3]
function Is-AltTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)[ab]\d*\]$')
}

# Identify bracket tags that are disk noise markers and should be ignored
function Is-DiskNoiseTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)\s*disks?\b')
}

# Identify bracket tags that mirror the extension and should be ignored
function Is-ExtensionNoiseTag {
    param(
        [Parameter(Mandatory=$true)][string]$Tag,
        [Parameter(Mandatory=$true)][string]$FileName
    )

    # Alt tags are never extension-noise
    if (Is-AltTag $Tag) { return $false }

    # Pull extension and normalize
    $ext = [System.IO.Path]::GetExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($ext)) { return $false }
    $extNoDot = $ext.TrimStart('.')
    if ([string]::IsNullOrWhiteSpace($extNoDot)) { return $false }

    # Extract inner tag contents
    $m = [regex]::Match($Tag, '^\[(?<X>[^\]]+)\]$')
    if (-not $m.Success) { return $false }
    $inner = $m.Groups['X'].Value.Trim()

    # Match case-insensitively vs extension
    return ($inner -ieq $extNoDot)
}

# Trim a base title prefix to a stable normalized stem
function Clean-BasePrefix {
    param([string]$Prefix)

    # Null-safe normalization
    if ($null -eq $Prefix) { return "" }

    # Trim and strip trailing punctuation artifacts
    $p = $Prefix.Trim()
    $p = $p -replace '[\s._-]+$', ''
    $p = $p -replace '\(\s*$', ''

    return $p.Trim()
}

# Print a phase banner line
function Write-Phase {
    param([string]$Text)
    Write-Host ""
    Write-Host $Text -ForegroundColor Cyan
}

# Build an alt fallback chain (e.g., [a2] -> [a] -> base)
function Get-AltFallbackChain {
    param([string]$AltKey)

    # Empty alt => base only
    if ([string]::IsNullOrWhiteSpace($AltKey)) { return @("") }

    # Parse TOSEC alt token
    $m = [regex]::Match($AltKey, '^\[(?i)(?<L>[ab])(?<N>\d*)\]$')
    if (-not $m.Success) { return @($AltKey, "") }

    $letter = $m.Groups['L'].Value.ToLowerInvariant()
    $num = $m.Groups['N'].Value

    # If no number, only itself and base
    if ([string]::IsNullOrWhiteSpace($num)) {
        return @($AltKey, "")
    }

    # If numbered alt, try itself, then the unnumbered, then base
    return @($AltKey, "[$letter]", "")
}

# Remove [!] from a tag key to create a non-bang compatibility key
function Get-NonBangTagsKey {
    param([string]$BaseTagsKey)
    if ([string]::IsNullOrWhiteSpace($BaseTagsKey)) { return "" }
    return ($BaseTagsKey -replace '\[\!\]', '')
}

# Normalize alt value to stable representation
function Normalize-Alt {
    param([AllowNull()][AllowEmptyString()][string]$Alt)
    if ([string]::IsNullOrWhiteSpace($Alt)) { return $null }
    return $Alt
}

# Determine platform root folder name relative to script location
function Get-PlatformRootName {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$ScriptDir
    )

    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $dirFull    = (Resolve-Path -LiteralPath $Directory).Path.TrimEnd('\')

    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # If script is in ROMs root, platform is first relative path segment
    if ($scriptIsRomsRoot) {
        if (-not $dirFull.StartsWith($scriptFull, [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
        $rel = $dirFull.Substring($scriptFull.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($rel)) { return $null }
        $parts = $rel -split '\\'
        if ($parts.Count -ge 1) { return $parts[0].ToLowerInvariant() }
        return $null
    }

    # If script is in a platform folder, that folder is the platform
    return $scriptLeaf.ToLowerInvariant()
}

# Resolve a platform root filesystem path
function Get-PlatformRootPath {
    param(
        [Parameter(Mandatory=$true)][string]$ScriptDir,
        [Parameter(Mandatory=$true)][string]$PlatformRootName
    )

    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # If script is in ROMs root, platform root path is ROMs\platform
    if ($scriptIsRomsRoot) {
        return (Join-Path $scriptFull $PlatformRootName)
    }

    # If script is in a platform folder, that folder is platform root
    return $scriptFull
}

# Produce a stable platform label for count reporting
function Get-PlatformCountLabel {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$ScriptDir
    )

    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $dirFull    = (Resolve-Path -LiteralPath $Directory).Path.TrimEnd('\')

    # If directory isn't under script root, fall back to leaf name
    if (-not $dirFull.StartsWith($scriptFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return (Split-Path -Leaf $dirFull).ToUpperInvariant()
    }

    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $rel = $dirFull.Substring($scriptFull.Length).TrimStart('\')
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($rel)) { $parts = $rel -split '\\' }

    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # If ROMs root, label as PLATFORM\subpath
    if ($scriptIsRomsRoot) {
        if ($parts.Count -eq 0) { return $scriptLeaf.ToUpperInvariant() }
        $platform = $parts[0].ToUpperInvariant()
        $subParts = if ($parts.Count -gt 1) { $parts[1..($parts.Count-1)] } else { @() }
        if ($subParts.Count -gt 0) { return ($platform + "\" + ($subParts -join "\")) }
        return $platform
    }

    # If platform root, label as PLATFORM\subpath
    $platform = $scriptLeaf.ToUpperInvariant()
    if ($parts.Count -gt 0) { return ($platform + "\" + ($parts -join "\")) }
    return $platform
}

# Convert an absolute file path to a ./ relative gamelist path
function Get-RelativeGamelistPath {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformRootPath,
        [Parameter(Mandatory=$true)][string]$FileFullPath
    )

    # Resolve paths safely
    try {
        $rootFull = (Resolve-Path -LiteralPath $PlatformRootPath).Path.TrimEnd('\')
        $fileFull = (Resolve-Path -LiteralPath $FileFullPath).Path
    } catch {
        return $null
    }

    # Ensure file is under platform root
    if (-not $fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    # Build ./relative path with forward slashes
    $rel = $fileFull.Substring($rootFull.Length).TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($rel)) { return $null }

    $rel = $rel -replace '\\', '/'
    return ("./" + $rel)
}

# Generate a unique gamelist.backup path (gamelist.backup, gamelist.backup (1), ...)
function Get-UniqueGamelistBackupPath {
    param([Parameter(Mandatory=$true)][string]$GamelistPath)

    $dir = Split-Path -Parent $GamelistPath
    $base = Join-Path $dir "gamelist.backup"

    # If base backup doesn't exist, use it
    if (-not (Test-Path -LiteralPath $base)) { return $base }

    # Otherwise find next available numbered backup name
    $i = 1
    while ($true) {
        $p = Join-Path $dir ("gamelist.backup ({0})" -f $i)
        if (-not (Test-Path -LiteralPath $p)) { return $p }
        $i++
    }
}

# Load gamelist.xml into cache for a platform
function Ensure-GamelistLoaded {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformRootLower,
        [Parameter(Mandatory=$true)][string]$PlatformRootPath
    )

    # Return cached state if present
    if ($gamelistStateByPlatform.ContainsKey($PlatformRootLower)) {
        return $gamelistStateByPlatform[$PlatformRootLower]
    }

    $gamelistPath = Join-Path $PlatformRootPath "gamelist.xml"

    # Create new state object
    $state = [PSCustomObject]@{
        RootPath      = $PlatformRootPath
        GamelistPath  = $gamelistPath
        Lines         = $null
        Changed       = $false
        Exists        = (Test-Path -LiteralPath $gamelistPath)
    }

    # If gamelist exists, load all lines
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

# Save cached gamelist state to disk if changed
function Save-GamelistIfChanged {
    param([Parameter(Mandatory=$true)]$State)

    # Skip if no write is needed/possible
    if (-not $State.Exists -or -not $State.Changed -or $null -eq $State.Lines) { return $false }

    # Respect dry run
    if ($dryRun) { return $false }

    # Make a backup once per gamelist per run
    if (-not $gamelistBackupDone.ContainsKey($State.GamelistPath)) {
        $backupPath = Get-UniqueGamelistBackupPath -GamelistPath $State.GamelistPath
        Copy-Item -LiteralPath $State.GamelistPath -Destination $backupPath -Force
        $gamelistBackupDone[$State.GamelistPath] = $true
    }

    # Write file as UTF-8 without BOM
    $text = ($State.Lines -join [Environment]::NewLine)
    [System.IO.File]::WriteAllText($State.GamelistPath, $text, [System.Text.UTF8Encoding]::new($false))
    return $true
}

# Fetch <name> value for a game entry identified by a specific <path>
function Get-GamelistNameByRelPath {
    param(
        [Parameter(Mandatory=$true)]$Lines,
        [Parameter(Mandatory=$true)][string]$RelPath
    )

    # Null-safety
    if ($null -eq $Lines -or [string]::IsNullOrWhiteSpace($RelPath)) { return $null }

    # Scan lines to find matching <path>
    for ($i = 0; $i -lt $Lines.Count; $i++) {

        $line = $Lines[$i]

        # Skip null lines
        if ($null -eq $line) { continue }

        # Match a <path> line
        $m = [regex]::Match($line, '^\s*<path>\s*(?<V>.*?)\s*</path>\s*$')
        if (-not $m.Success) { continue }

        # Only proceed on exact path match
        if ($m.Groups['V'].Value -ine $RelPath) { continue }

        # Walk forward inside the <game> block to find <name>
        $j = $i + 1
        while ($j -lt $Lines.Count) {
            $tline = $Lines[$j]
            if ($tline -match '^\s*<path>\s*') { break }
            if ($tline -match '^\s*</game>\s*$') { break }

            $nm = [regex]::Match($tline, '^\s*<name>\s*(?<N>.*?)\s*</name>\s*$')
            if ($nm.Success) {
                $val = $nm.Groups['N'].Value
                if (-not [string]::IsNullOrWhiteSpace($val)) { return $val }
                return ""
            }

            $j++
        }

        return $null
    }

    return $null
}

# Build a fast lookup set of all <path> values present in a gamelist.xml line array
function Get-GamelistPathIndex {
    param(
        [Parameter(Mandatory=$true)]$Lines
    )

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ($null -eq $Lines) { return $set }

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        if ($null -eq $line) { continue }

        $m = [regex]::Match($line, '^\s*<path>\s*(?<V>.*?)\s*</path>\s*$')
        if ($m.Success) {
            $v = $m.Groups['V'].Value
            if (-not [string]::IsNullOrWhiteSpace($v)) {
                [void]$set.Add($v)
            }
        }
    }

    return $set
}

# Returns $true if the file is inside any folder whose name is in $skipFolders (at any depth)
function Is-InSkipFolderPath {
    param(
        [Parameter(Mandatory=$true)][string]$FileFullPath,
        [Parameter(Mandatory=$true)][string[]]$SkipFolderNames
    )

    if ([string]::IsNullOrWhiteSpace($FileFullPath)) { return $false }

    $di = [System.IO.DirectoryInfo]::new([System.IO.Path]::GetDirectoryName($FileFullPath))

    while ($null -ne $di) {
        if ($SkipFolderNames -contains $di.Name.ToLowerInvariant()) { return $true }
        $di = $di.Parent
    }

    return $false
}

# Ensure <name> appears above <hidden> within the same <game> block for a given <path> line index
function Ensure-NameBeforeHiddenInGameBlock {
    param(
        [Parameter(Mandatory=$true)][ref]$Lines,
        [Parameter(Mandatory=$true)][int]$PathLineIndex
    )

    $arr = $Lines.Value
    if ($null -eq $arr) { return $false }
    if ($PathLineIndex -lt 0 -or $PathLineIndex -ge $arr.Count) { return $false }

    # ---- Safe block boundary detection + hard scan cap ----
    $maxScan = 300
    $endIndex = $null
    $kStop = [Math]::Min($arr.Count - 1, $PathLineIndex + $maxScan)

    for ($k = $PathLineIndex; $k -le $kStop; $k++) {
        if ($arr[$k] -match '^\s*</game>\s*$') {
            $endIndex = $k
            break
        }
    }

    # If we couldn't find </game> within the safety window, do nothing
    if ($null -eq $endIndex) { return $false }
    # ----------------------------------------------------------------------

    # Locate first <name> and first <hidden> inside this block
    $nameIndex = $null
    $hiddenIndex = $null

    for ($k = $PathLineIndex + 1; $k -le $endIndex -and $k -lt $arr.Count; $k++) {

        $line = $arr[$k]
        if ($null -eq $line) { continue }

        if ($null -eq $nameIndex -and $line -match '^\s*<name>\s*.*?\s*</name>\s*$') {
            $nameIndex = $k
            continue
        }

        if ($null -eq $hiddenIndex -and $line -match '^\s*<hidden>\s*.*?\s*</hidden>\s*$') {
            $hiddenIndex = $k
            continue
        }

        if ($null -ne $nameIndex -and $null -ne $hiddenIndex) { break }
    }

    # If both exist and hidden is above name, swap the two lines
    if ($null -ne $nameIndex -and $null -ne $hiddenIndex -and $hiddenIndex -lt $nameIndex) {

        $tmp = $arr[$hiddenIndex]
        $arr[$hiddenIndex] = $arr[$nameIndex]
        $arr[$nameIndex] = $tmp

        $Lines.Value = $arr
        return $true
    }

    return $false
}

# Ensure every target disk entry has a canonical <name> line immediately after <path>
function Ensure-GameNamesInGamelist {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)]$Targets,
        [Parameter(Mandatory=$true)][string]$FallbackName
    )

    # Skip if gamelist isn't loaded
    if (-not $State.Exists -or $null -eq $State.Lines) { return $false }

    # ---- SMART IMPROVEMENT (B): use List[string] for efficient inserts/edits ----
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $State.Lines) { [void]$lines.Add($ln) }
    # ---------------------------------------------------------------------------

    # Identify Disk 1 target and prefer its <name> as canonical
    $disk1Target = $null
    foreach ($t in $Targets) {
        if ($t.DiskSort -eq 1) { $disk1Target = $t; break }
    }

    # Choose canonical name: Disk 1 <name> wins; else fallback
    $canonical = $null
    if ($null -ne $disk1Target -and (-not [string]::IsNullOrWhiteSpace($disk1Target.RelPath))) {
        $disk1Name = Get-GamelistNameByRelPath -Lines $lines -RelPath $disk1Target.RelPath
        if (-not [string]::IsNullOrWhiteSpace($disk1Name)) { $canonical = $disk1Name }
    }

    if ([string]::IsNullOrWhiteSpace($canonical)) { $canonical = $FallbackName }
    if ([string]::IsNullOrWhiteSpace($canonical)) { return $false }

    # Fast skip if all targets already match canonical name
    $allMatch = $true
    foreach ($t in $Targets) {
        if ([string]::IsNullOrWhiteSpace($t.RelPath)) { $allMatch = $false; break }
        $nm = Get-GamelistNameByRelPath -Lines $lines -RelPath $t.RelPath
        # ---- SMART IMPROVEMENT (A): explicitly treat missing name as mismatch ----
        if ($null -eq $nm -or $nm -ne $canonical) { $allMatch = $false; break }
        # -----------------------------------------------------------------------
    }
    if ($allMatch) { return $false }

    $didWork = $false

    # Enforce canonical <name> per target (insert or replace)
    foreach ($t in $Targets) {

        $rel = $t.RelPath
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }

        # Scan file to locate this entry's <path> line
        for ($i = 0; $i -lt $lines.Count; $i++) {

            $line = $lines[$i]
            if ($null -eq $line) { continue }

            $m = [regex]::Match($line, '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
            if (-not $m.Success) { continue }
            if ($m.Groups['V'].Value -ine $rel) { continue }

            $indent = $m.Groups['I'].Value

            # Walk forward inside block to find existing <name> line
            $j = $i + 1
            $nameLineIndex = $null
            $nameValue = $null

            while ($j -lt $lines.Count) {

                $tline = $lines[$j]
                if ($tline -match '^\s*<path>\s*') { break }
                if ($tline -match '^\s*</game>\s*$') { break }

                $nm = [regex]::Match($tline, '^\s*<name>\s*(?<N>.*?)\s*</name>\s*$')
                if ($nm.Success) {
                    $nameLineIndex = $j
                    $nameValue = $nm.Groups['N'].Value
                    break
                }

                $j++
            }

            # If name exists, replace if different from canonical
            if ($null -ne $nameLineIndex) {

                if ($nameValue -ne $canonical) {
                    if (-not $dryRun) {
                        $lines[$nameLineIndex] = ($indent + "<name>$canonical</name>")
                    }
                    $State.Changed = $true
                    $didWork = $true
                }

                # Ensure <name> is positioned above <hidden> within this <game> block
                $tmpArr = $lines.ToArray()
                if (Ensure-NameBeforeHiddenInGameBlock -Lines ([ref]$tmpArr) -PathLineIndex $i) {
                    if (-not $dryRun) {
                        $lines.Clear()
                        foreach ($ln in $tmpArr) { [void]$lines.Add($ln) }
                    }
                    $State.Changed = $true
                    $didWork = $true
                }

            } else {

                # If name doesn't exist, insert immediately after <path>
                if (-not $dryRun) {
                    $insertLine = ($indent + "<name>$canonical</name>")
                    $lines.Insert($i + 1, $insertLine)
                }

                $State.Changed = $true
                $didWork = $true

                # Ensure <name> is positioned above <hidden> within this <game> block
                $tmpArr = $lines.ToArray()
                if (Ensure-NameBeforeHiddenInGameBlock -Lines ([ref]$tmpArr) -PathLineIndex $i) {
                    if (-not $dryRun) {
                        $lines.Clear()
                        foreach ($ln in $tmpArr) { [void]$lines.Add($ln) }
                    }
                    $State.Changed = $true
                    $didWork = $true
                }

                $i++
            }

            break
        }
    }

    $State.Lines = $lines.ToArray()
    return $didWork
}

# Hide target entries in gamelist.xml by ensuring <hidden>true</hidden>
function Hide-GameEntriesInGamelist {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)]$Targets,
        [Parameter(Mandatory=$true)][hashtable]$UsedFiles,
        [switch]$SuppressAlreadyReport
    )

    $result = [PSCustomObject]@{
        DidWork            = $false
        NewlyHiddenCount   = 0
        AlreadyHiddenCount = 0
        MissingCount       = 0
    }

    # If gamelist isn't present/loaded, mark all targets missing
    if (-not $State.Exists -or $null -eq $State.Lines) {

        foreach ($t in $Targets) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true

            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
            })

            $result.MissingCount++
        }

        return $result
    }

    # ---- SMART IMPROVEMENT (B): use List[string] for efficient inserts/edits ----
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $State.Lines) { [void]$lines.Add($ln) }
    # ---------------------------------------------------------------------------

    # For each target, locate it and enforce <hidden>true</hidden>
    foreach ($t in $Targets) {

        $rel = $t.RelPath

        # If no relative path, treat as missing entry
        if ([string]::IsNullOrWhiteSpace($rel)) {

            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
            continue
        }

        $found = $false
        $handled = $false

        # Scan gamelist lines to locate this <path>
        for ($i = 0; $i -lt $lines.Count; $i++) {

            $line = $lines[$i]
            if ($null -eq $line) { continue }

            $m = [regex]::Match($line, '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
            if (-not $m.Success) { continue }

            $val = $m.Groups['V'].Value
            if ($val -ine $rel) { continue }

            $found = $true
            $indent = $m.Groups['I'].Value

            # Walk within this <game> block to find existing <hidden>
            $j = $i + 1
            $hiddenLineIndex = $null
            $hiddenValue = $null

            while ($j -lt $lines.Count) {
                $tline = $lines[$j]

                if ($tline -match '^\s*<path>\s*') { break }
                if ($tline -match '^\s*</game>\s*$') { break }

                $hm = [regex]::Match($tline, '^\s*<hidden>\s*(?<H>.*?)\s*</hidden>\s*$')
                if ($hm.Success) {
                    $hiddenLineIndex = $j
                    $hiddenValue = $hm.Groups['H'].Value
                    break
                }
                $j++
            }

            # If <hidden> exists, ensure it is true; else insert a new line
            if ($null -ne $hiddenLineIndex) {

                if ($hiddenValue -match '^(?i)true$') {
                    $UsedFiles[$t.FullPath] = $true

                    if (-not $SuppressAlreadyReport) {

                        # Micro dedupe safety-net (by FullPath)
                        if ($noM3UAlreadyHiddenSet.Add([string]$t.FullPath)) {
                            [void]$noM3UAlreadyHidden.Add([PSCustomObject]@{
                                FullPath = $t.FullPath
                                Reason   = "Already hidden in gamelist.xml"
                            })
                            $result.AlreadyHiddenCount++
                        }
                    }

                    # Ensure <name> is positioned above <hidden> within this <game> block
                    $tmpArr = $lines.ToArray()
                    if (Ensure-NameBeforeHiddenInGameBlock -Lines ([ref]$tmpArr) -PathLineIndex $i) {
                        if (-not $dryRun) {
                            $lines.Clear()
                            foreach ($ln in $tmpArr) { [void]$lines.Add($ln) }
                        }
                        $State.Changed = $true
                        $result.DidWork = $true
                    }

                    $handled = $true
                }
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

                    # Ensure <name> is positioned above <hidden> within this <game> block
                    $tmpArr = $lines.ToArray()
                    if (Ensure-NameBeforeHiddenInGameBlock -Lines ([ref]$tmpArr) -PathLineIndex $i) {
                        if (-not $dryRun) {
                            $lines.Clear()
                            foreach ($ln in $tmpArr) { [void]$lines.Add($ln) }
                        }
                        $State.Changed = $true
                        $result.DidWork = $true
                    }

                    $handled = $true
                }
            }
            else {

                if (-not $dryRun) {
                    $insertLine = ($indent + "<hidden>true</hidden>")
                    $lines.Insert($i + 1, $insertLine)
                }

                $State.Changed = $true
                $result.DidWork = $true
                $result.NewlyHiddenCount++
                $UsedFiles[$t.FullPath] = $true
                [void]$noM3UNewlyHidden.Add([PSCustomObject]@{
                    FullPath = $t.FullPath
                    Reason   = "Hidden in gamelist.xml"
                })

                # Ensure <name> is positioned above <hidden> within this <game> block
                $tmpArr = $lines.ToArray()
                if (Ensure-NameBeforeHiddenInGameBlock -Lines ([ref]$tmpArr) -PathLineIndex $i) {
                    if (-not $dryRun) {
                        $lines.Clear()
                        foreach ($ln in $tmpArr) { [void]$lines.Add($ln) }
                    }
                    $State.Changed = $true
                    $result.DidWork = $true
                }

                $handled = $true
                $i++
            }

            if ($handled) { break }
        }

        # If we never found the entry, record as missing
        if (-not $found) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
        }
    }

    $State.Lines = $lines.ToArray()
    return $result
}

# Unhide target entries in gamelist.xml by removing <hidden>true</hidden>
function Unhide-GameEntriesInGamelist {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)]$Targets,
        [Parameter(Mandatory=$true)][hashtable]$UsedFiles,
        [switch]$SuppressAlreadyReport
    )

    $result = [PSCustomObject]@{
        DidWork             = $false
        NewlyUnhiddenCount  = 0
        AlreadyVisibleCount = 0
        MissingCount        = 0
    }

    # If gamelist isn't present/loaded, mark all targets missing
    if (-not $State.Exists -or $null -eq $State.Lines) {

        foreach ($t in $Targets) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
        }

        return $result
    }

    # ---- SMART IMPROVEMENT (B): use List[string] for efficient deletes/edits ----
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $State.Lines) { [void]$lines.Add($ln) }
    # ---------------------------------------------------------------------------

    # For each target, locate and remove <hidden>true</hidden> if present
    foreach ($t in $Targets) {

        $rel = $t.RelPath

        # If no relative path, treat as missing entry
        if ([string]::IsNullOrWhiteSpace($rel)) {

            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
            continue
        }

        $found = $false
        $handled = $false

        # Scan gamelist lines to locate this <path>
        for ($i = 0; $i -lt $lines.Count; $i++) {

            $line = $lines[$i]
            if ($null -eq $line) { continue }

            $m = [regex]::Match($line, '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
            if (-not $m.Success) { continue }

            $val = $m.Groups['V'].Value
            if ($val -ine $rel) { continue }

            $found = $true

            # Ensure <name> is positioned above <hidden> within this <game> block (formatting enforcement)
            $tmpArr = $lines.ToArray()
            if (Ensure-NameBeforeHiddenInGameBlock -Lines ([ref]$tmpArr) -PathLineIndex $i) {
                if (-not $dryRun) {
                    $lines.Clear()
                    foreach ($ln in $tmpArr) { [void]$lines.Add($ln) }
                }
                $State.Changed = $true
                $result.DidWork = $true
            }

            # Walk within this <game> block to find existing <hidden>
            $j = $i + 1
            $hiddenLineIndex = $null
            $hiddenValue = $null

            while ($j -lt $lines.Count) {
                $tline = $lines[$j]

                if ($tline -match '^\s*<path>\s*') { break }
                if ($tline -match '^\s*</game>\s*$') { break }

                $hm = [regex]::Match($tline, '^\s*<hidden>\s*(?<H>.*?)\s*</hidden>\s*$')
                if ($hm.Success) {
                    $hiddenLineIndex = $j
                    $hiddenValue = $hm.Groups['H'].Value
                    break
                }

                $j++
            }

            # If hidden=true, remove that line; else record already visible
            if ($null -ne $hiddenLineIndex -and $hiddenValue -match '^(?i)true$') {

                if (-not $dryRun) {
                    $lines.RemoveAt($hiddenLineIndex)
                }

                $State.Changed = $true
                $result.DidWork = $true
                $result.NewlyUnhiddenCount++
                $UsedFiles[$t.FullPath] = $true

                [void]$noM3UNewlyUnhidden.Add([PSCustomObject]@{
                    FullPath = $t.FullPath
                    Reason   = "Unhidden in gamelist.xml"
                })

                $handled = $true
                break
            }
            else {

                $UsedFiles[$t.FullPath] = $true

                if (-not $SuppressAlreadyReport) {

                    # Micro dedupe safety-net (by FullPath)
                    if ($noM3UAlreadyVisibleSet.Add([string]$t.FullPath)) {
                        [void]$noM3UAlreadyVisible.Add([PSCustomObject]@{
                            FullPath = $t.FullPath
                            Reason   = "Already visible in gamelist.xml"
                        })
                        $result.AlreadyVisibleCount++
                    }
                }

                $handled = $true
                break
            }
        }

        # If we never found the entry, record as missing
        if (-not $found) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
        }
    }

    $State.Lines = $lines.ToArray()
    return $result
}

# Parse a candidate filename into base title, disk/side markers, and tag keys
function Parse-GameFile {
    param(
        [Parameter(Mandatory=$true)][string]$FileName,
        [Parameter(Mandatory=$true)][string]$Directory
    )

    # Strip extension for token parsing
    $nameNoExt = $FileName -replace '\.[^\.]+$', ''

    # Define disk/disc and side-only regex patterns
    $diskPattern = '(?i)^(?<Prefix>.*?)(?<Sep>[\s._\-\(]|^)(?<Type>disk|disc)(?!s)[\s_]*(?<Disk>\d+|[A-Za-z]|(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX))(?:(?=\s+of\s+\d+)|(?=\s+Side\s+[A-Za-z])|(?=[\s\)\]\._-]|$))(?:\s+of\s+(?<Total>\d+))?(?:\s+Side\s+(?<Side>[A-Za-z]))?(?<After>.*)$'
    $sideOnlyPattern = '(?i)^(?<Prefix>.*?)(?<Sep>[\s._\-\(]|^)Side\s+(?<SideOnly>[A-Za-z])(?=[\s\)\]\._-]|$)(?<After>.*)$'

    # Attempt disk-pattern match first
    $diskMatch = [regex]::Match($nameNoExt, $diskPattern)
    $sideOnlyMatch = $null

    $hasDisk = $diskMatch.Success

    # If no disk match, attempt side-only match
    if (-not $hasDisk) {
        $sideOnlyMatch = [regex]::Match($nameNoExt, $sideOnlyPattern)
    }

    # If neither match, not a multi-disk candidate
    if (-not $hasDisk -and (-not $sideOnlyMatch.Success)) {
        return $null
    }

    # Initialize parsed token fields
    $prefixRaw  = ""
    $diskToken  = $null
    $totalToken = $null
    $sideToken  = $null
    $after      = ""

    # Populate fields from disk match or side-only match
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

    # Normalize base prefix and post-token text
    $basePrefix = Clean-BasePrefix $prefixRaw
    $afterNorm = $after -replace '^[\)\s]+', ''

    # Extract a shared name hint if present as leading parenthetical
    $nameHint = ""
    $beforeBracket = $afterNorm
    $bracketIdx = $beforeBracket.IndexOf('[')
    if ($bracketIdx -ge 0) { $beforeBracket = $beforeBracket.Substring(0, $bracketIdx) }
    $mHint = [regex]::Match($beforeBracket, '^\s*(\([^\)]+\))')
    if ($mHint.Success) { $nameHint = $mHint.Groups[1].Value }

    # Extract bracket tags while filtering disk-noise and extension-noise
    $bracketTags = @()
    foreach ($m in [regex]::Matches($afterNorm, '\[[^\]]+\]')) {
        $tag = $m.Value
        if (Is-DiskNoiseTag $tag) { continue }
        if (Is-ExtensionNoiseTag -Tag $tag -FileName $FileName) { continue }
        $bracketTags += $tag
    }

    # Partition tags into alt tag vs base tags
    $altTag  = ""
    $baseTags = @()
    foreach ($t in $bracketTags) {
        if (Is-AltTag $t) { $altTag = $t }
        else { $baseTags += $t }
    }

    $baseTagsKey = ($baseTags -join "")

    # Convert tokens to sort keys
    $diskSort = Convert-DiskToSort $diskToken
    $sideSort = Convert-SideToSort $sideToken

    # Parse total disks token if numeric
    $totalDisks = $null
    if ($totalToken -match '^\d+$') { $totalDisks = [int]$totalToken }

    # Return parsed object
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

# Select candidate entries for a disk number and alt key under total constraints
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

    # Normalize requested alt
    $wantAlt = Normalize-Alt $AltTag

    # Filter candidates by disk number, alt, and total constraints
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

# Initialize parsed candidate collection
$parsed = @()

# Display scan banner
Write-Phase "Collecting ROM file data (scanning folders)..."

# Enumerate candidate ROM files and parse disk/side patterns
Get-ChildItem -Path $scriptDir -File -Recurse -Depth 2 | ForEach-Object {

    # Skip M3U files
    if ($_.Extension -ieq ".m3u") { return }

    # Skip known media/manual folders (match any ancestor folder name)
    if (Is-InSkipFolderPath -FileFullPath $_.FullName -SkipFolderNames $skipFolders) { return }

    # If NON-M3U mode is skip, reject files under NON-M3U platforms
    if ($noM3UPlatformMode -ieq "skip") {

        $plat = Get-PlatformRootName -Directory $_.DirectoryName -ScriptDir $scriptDir
        if ($null -ne $plat) {

            $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
            if ($noM3USetLower -contains $plat.ToLowerInvariant()) { return }
        }
    }

    # Parse the file name into a multi-disk candidate if applicable
    $p = Parse-GameFile -FileName $_.Name -Directory $_.DirectoryName
    if ($null -ne $p) { $parsed += $p }
}

# ==================================================================================================
# PHASE 2: INDEXING / GROUPING
# ==================================================================================================

# Display indexing banner
Write-Phase "Indexing parsed candidates (grouping titles, tags, variants, etc.)..."

# Build strict groups by directory + base prefix + base tags
$groupsStrict = $parsed | Group-Object Directory, BasePrefix, BaseTagsKey

# Build titleKey -> list index (for relaxed grouping passes)
$titleIndex = @{}
foreach ($p in $parsed) {
    if (-not $titleIndex.ContainsKey($p.TitleKey)) { $titleIndex[$p.TitleKey] = @() }
    $titleIndex[$p.TitleKey] += $p
}

# Build [!] presence indices for [!] suppression rules
$bangByTitleNB = @{}
$bangAltByTitleNB = @{}

foreach ($p in $parsed) {

    if ($p.BaseTagsKey -match '\[\!\]') {

        $k = $p.TitleKey + "`0" + $p.BaseTagsKeyNB
        $bangByTitleNB[$k] = $true

        if (-not [string]::IsNullOrWhiteSpace($p.AltTag)) {
            $k2 = $p.TitleKey + "`0" + $p.BaseTagsKeyNB + "`0" + $p.AltTag
            $bangAltByTitleNB[$k2] = $true
        }
    }
}

# Initialize collision and signature trackers
$occupiedPlaylistPaths = @{}
$m3uWrittenPlaylistPaths = @{}
$playlistSignatures = @{}

$suppressedDuplicatePlaylists   = @{} # playlistPath -> collidedWithPath
$suppressedPreExistingPlaylists = @{} # playlistPath -> $true (content identical)
$overwrittenExistingPlaylists   = @{} # playlistPath -> $true (content differed)

# Initialize used-files tracker
$usedFiles = @{}

# ==================================================================================================
# PHASE 3: PROCESS MULTI-DISK GROUPS (M3U PLAYLISTS + NON-M3U GAMELIST HIDING)
# ==================================================================================================

# Display processing banner
Write-Phase "Processing multi-disk candidates (playlists / gamelist updates)..."

# Iterate each strict group (dir + base title + base tags)
foreach ($group in $groupsStrict) {

    $groupFiles = $group.Group
    $directory  = $groupFiles[0].Directory
    $basePrefix = $groupFiles[0].BasePrefix
    $titleKey   = $groupFiles[0].TitleKey

    # Resolve title-wide compatible file set for NB-key comparisons
    $titleFiles = if ($titleIndex.ContainsKey($titleKey)) { $titleIndex[$titleKey] } else { $groupFiles }
    $strictNBKey = $groupFiles[0].BaseTagsKeyNB

    $titleCompatible = @(
        $titleFiles | Where-Object { $_.BaseTagsKeyNB -eq $strictNBKey }
    )

    # Determine alt keys present in this group
    $altKeys = @(
        ($groupFiles | Select-Object -ExpandProperty AltTag | ForEach-Object { if ($_ -eq $null) { "" } else { $_ } } | Sort-Object -Unique)
    )

    # Iterate alt variants
    foreach ($altKey in $altKeys) {

        # Identify Disk 1 roots and potential total counts
        $disk1Roots = @($titleCompatible | Where-Object { $_.DiskSort -eq 1 })

        $totals = @($disk1Roots | Where-Object { $_.TotalDisks -ne $null } | Select-Object -ExpandProperty TotalDisks | Sort-Object -Unique)
        if ($totals.Count -eq 0) { $totals = @($null) }

        # Iterate total-count variants (or null total)
        foreach ($rootTotal in $totals) {

            # Choose expected disk targets based on declared total or observed disks
            $diskTargets = if ($rootTotal) { 1..$rootTotal }
                           else { @($titleCompatible | Select-Object -ExpandProperty DiskSort | Sort-Object -Unique) }

            $playlistFiles = @()

            # Fill each disk target via selection passes
            foreach ($d in $diskTargets) {

                $picked = @()

                # Pass 1: strict within group
                $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $altKey -RootTotal $rootTotal

                # Pass 1b: allow single unambiguous alt disk when base is requested
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

                # Pass 2: relaxed by non-bang compatible title set
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

                # Pass 3: alt fallback chain within group
                if ($picked.Count -eq 0) {
                    $altChain = Get-AltFallbackChain $altKey
                    foreach ($tryAlt in $altChain) {
                        $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $tryAlt -RootTotal $rootTotal
                        if ($picked.Count -gt 0) { break }
                    }
                }

                # Pass 4: alt fallback chain within title-compatible
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

                # If found, append picks to playlist file set
                if ($picked.Count -gt 0) { $playlistFiles += $picked }
            }

            # Require at least two files to qualify as a set
            if (@($playlistFiles).Count -lt 2) { continue }

            # Determine optional shared name hint
            $uniqueHints = @($playlistFiles | Select-Object -ExpandProperty NameHint | Sort-Object -Unique)
            $useHint = ""
            if ($uniqueHints.Count -eq 1 -and (-not [string]::IsNullOrWhiteSpace($uniqueHints[0]))) { $useHint = $uniqueHints[0] }

            # Build base playlist name (title + optional hint)
            $playlistBase = $basePrefix
            if ($useHint) { $playlistBase += $useHint }

            # Compute tags common to all chosen entries for naming stability
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

            # Append alt tag if all selected entries share the same alt
            $altsInPlaylist = @($playlistFiles | ForEach-Object { Normalize-Alt $_.AltTag } | Sort-Object -Unique)
            if ($altsInPlaylist.Count -eq 1 -and $altsInPlaylist[0]) {
                $playlistBase += $altsInPlaylist[0]
            }

            # Cleanup playlist base name
            $playlistBase = $playlistBase -replace '\s{2,}', ' '
            $playlistBase = $playlistBase -replace '[\s._-]*[\(]*$', ''
            $playlistBase = $playlistBase -replace '\(\s*\)', ''
            $playlistBase = $playlistBase.Trim()
            if ([string]::IsNullOrWhiteSpace($playlistBase)) { continue }

            # Determine whether this platform is NON-M3U
            $platformRoot = Get-PlatformRootName -Directory $directory -ScriptDir $scriptDir

            $isNoM3U = $false
            if ($null -ne $platformRoot) {
                $rootLower = $platformRoot.ToLowerInvariant()
                $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
                $isNoM3U = ($noM3USetLower -contains $rootLower)
            }

            # Respect "skip" mode for NON-M3U platforms
            if ($isNoM3U -and $noM3UPlatformMode -ieq "skip") { continue }

            # Track encountered NON-M3U platforms for Phase 3.5
            if ($isNoM3U -and $null -ne $platformRoot) {
                $noM3UPlatformsEncountered[$platformRoot.ToLowerInvariant()] = $true
            }

            # Sort final set entries into disk/side order
            $sorted = $playlistFiles | Sort-Object DiskSort, SideSort

            # Enforce Disk 1 requirement in NON-M3U mode (or treat as incomplete)
            if ($isNoM3U) {

                $disk1Candidates = @($sorted | Where-Object { $_.DiskSort -eq 1 } | Sort-Object SideSort)
                $hasDisk1 = ($disk1Candidates.Count -gt 0)

                if (-not $hasDisk1) {

                    $platformLower = $platformRoot.ToLowerInvariant()
                    $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
                    $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

                    $targetsAll = @()
                    foreach ($sf in $sorted) {
                        $fullFile = Join-Path $sf.Directory $sf.FileName
                        $rel = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $fullFile
                        $targetsAll += [PSCustomObject]@{
                            FullPath      = $fullFile
                            RelPath       = $rel
                            PlatformLabel = $platformRoot.ToUpperInvariant()
                            DiskSort      = $sf.DiskSort
                            SideSort      = $sf.SideSort
                        }
                    }

                    # Ensures consistent <name> lines for NON-M3U disk entries when working with gamelist.xml
                    Ensure-GameNamesInGamelist -State $state -Targets $targetsAll -FallbackName $playlistBase | Out-Null

                    # Unhide everything when Disk 1 is missing (safety)
                    $unhideResult = Unhide-GameEntriesInGamelist -State $state -Targets $targetsAll -UsedFiles $usedFiles

                    $platLabel = $platformRoot.ToUpperInvariant()
                    if (-not $gamelistUnhiddenCounts.ContainsKey($platLabel)) { $gamelistUnhiddenCounts[$platLabel] = 0 }
                    if (-not $gamelistAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistAlreadyVisibleCounts[$platLabel] = 0 }
                    $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
                    $gamelistAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount

                    foreach ($sf in $sorted) {
                        $full = Join-Path $sf.Directory $sf.FileName
                        [void]$noM3UNoDisk1Sets.Add([PSCustomObject]@{
                            FullPath = $full
                            Reason   = "Disk 1 not detected; set treated as incomplete"
                        })
                        $usedFiles[$full] = $true
                    }

                    continue
                }
            }

            # Determine completeness when total is declared
            $setIsIncomplete = $false
            if ($rootTotal) {
                $presentDisks = @($playlistFiles | Select-Object -ExpandProperty DiskSort | Sort-Object -Unique)
                $missingDisksLocal = @()
                foreach ($ed in (1..$rootTotal)) {
                    if (-not ($presentDisks -contains $ed)) { $missingDisksLocal += $ed }
                }
                if ($missingDisksLocal.Count -gt 0) { $setIsIncomplete = $true }
            }

            # Apply [!] suppression rule when appropriate
            $thisHasBang = ($groupFiles[0].BaseTagsKey -match '\[\!\]')
            $wantAltNorm = Normalize-Alt $altKey
            if (-not $thisHasBang -and $wantAltNorm) {
                $kBang    = $titleKey + "`0" + $strictNBKey
                $kBangAlt = $titleKey + "`0" + $strictNBKey + "`0" + $wantAltNorm
                if ($bangByTitleNB.ContainsKey($kBang) -and $bangAltByTitleNB.ContainsKey($kBangAlt)) {
                    continue
                }
            }

            # Build cleaned M3U content lines
            $newLines = @($sorted | ForEach-Object { $_.FileName })
            $cleanLines = @(
                $newLines |
                    ForEach-Object { $_.TrimEnd() } |
                    Where-Object { $_ -ne "" }
            )

            $cleanText = ($cleanLines -join "`n")
            $newNorm   = Normalize-M3UText $cleanText

            # Choose playlist output path (M3U vs NON-M3U sentinel)
            # ---- NON-M3U uses internal sentinel id (not a fake filesystem path) ----
            $playlistPath = if ($isNoM3U) {
                ("NONM3U::{0}::{1}::{2}::{3}::{4}" -f $platformRoot, $titleKey, $strictNBKey, $wantAltNorm, $rootTotal)
            } else {
                (Join-Path $directory "$playlistBase.m3u")
            }
            # --------------------------------------------------------------------------------------

            # Avoid same-run path collisions via [alt] suffix
            if ($occupiedPlaylistPaths.ContainsKey($playlistPath)) {
                $altIndex = 1
                do {
                    $suffix = if ($altIndex -eq 1) { "[alt]" } else { "[alt$altIndex]" }
                    if ($isNoM3U) {
                        $playlistPath = ("NONM3U::{0}::{1}::{2}::{3}::{4}{5}" -f $platformRoot, $titleKey, $strictNBKey, $wantAltNorm, $rootTotal, $suffix)
                    } else {
                        $playlistPath = Join-Path $directory "$playlistBase$suffix.m3u"
                    }
                    $altIndex++
                } while ($occupiedPlaylistPaths.ContainsKey($playlistPath))
            }

            # Build a signature used to suppress duplicates within a run
            $sigParts = @()
            foreach ($sf in $sorted) { $sigParts += (Join-Path $sf.Directory $sf.FileName) }
            $playlistSig = ($sigParts -join "`0")

            # Suppress duplicate content playlists during this run
            if ($playlistSignatures.ContainsKey($playlistSig)) {

                $suppressedDuplicatePlaylists[$playlistPath] = $playlistSignatures[$playlistSig]
                $occupiedPlaylistPaths[$playlistPath] = $true

                foreach ($sf in $sorted) {
                    $full = Join-Path $sf.Directory $sf.FileName
                    $usedFiles[$full] = $true
                }

                continue
            }

            # For M3U paths, compare against existing file content to suppress or overwrite
            if (-not $isNoM3U) {

                $existsOnDisk = Test-Path -LiteralPath $playlistPath
                if ($existsOnDisk) {

                    $existingRaw = $null
                    try {
                        $existingRaw = Get-Content -LiteralPath $playlistPath -Raw -ErrorAction Stop
                    } catch {
                        $existingRaw = $null
                    }

                    $existingNorm = Normalize-M3UText $existingRaw

                    $sameContent = ($existingNorm.Count -eq $newNorm.Count)
                    if ($sameContent) {
                        for ($i = 0; $i -lt $existingNorm.Count; $i++) {
                            if ($existingNorm[$i] -ne $newNorm[$i]) { $sameContent = $false; break }
                        }
                    }

                    if ($sameContent) {
                        $suppressedPreExistingPlaylists[$playlistPath] = $true
                        $occupiedPlaylistPaths[$playlistPath] = $true
                        $playlistSignatures[$playlistSig] = $playlistPath

                        foreach ($sf in $sorted) {
                            $full = Join-Path $sf.Directory $sf.FileName
                            $usedFiles[$full] = $true
                        }

                        continue
                    }

                    $overwrittenExistingPlaylists[$playlistPath] = $true
                }
            }

            # Claim path and signature for this run
            $playlistSignatures[$playlistSig] = $playlistPath
            $occupiedPlaylistPaths[$playlistPath] = $true

            # NON-M3U branch: hide Disk 2+ in gamelist.xml (unless incomplete)
            if ($isNoM3U) {

                $disk1Candidates = @($sorted | Where-Object { $_.DiskSort -eq 1 } | Sort-Object SideSort)
                if ($disk1Candidates.Count -eq 0) { continue }

                $primaryObj  = $disk1Candidates[0]
                $primaryFull = Join-Path $primaryObj.Directory $primaryObj.FileName
                $usedFiles[$primaryFull] = $true

                # Resolve platform gamelist state first (rootPath/state are required below)
                $platformLower = $platformRoot.ToLowerInvariant()
                $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
                $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

                # Ensure primary is not hidden (safety: primary must be visible)
                $primaryRel = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $primaryFull
                $primaryTarget = @([PSCustomObject]@{
                    FullPath      = $primaryFull
                    RelPath       = $primaryRel
                    PlatformLabel = $platformRoot.ToUpperInvariant()
                    DiskSort      = 1
                    SideSort      = $primaryObj.SideSort
                })

                $unhidePrimaryResult = Unhide-GameEntriesInGamelist -State $state -Targets $primaryTarget -UsedFiles $usedFiles

                $platLabel = $platformRoot.ToUpperInvariant()
                if (-not $gamelistUnhiddenCounts.ContainsKey($platLabel)) { $gamelistUnhiddenCounts[$platLabel] = 0 }
                if (-not $gamelistAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistAlreadyVisibleCounts[$platLabel] = 0 }
                $gamelistUnhiddenCounts[$platLabel] += [int]$unhidePrimaryResult.NewlyUnhiddenCount
                $gamelistAlreadyVisibleCounts[$platLabel] += [int]$unhidePrimaryResult.AlreadyVisibleCount

                $platformLower = $platformRoot.ToLowerInvariant()
                $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
                $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

                $targetsAll = @()
                foreach ($sf in $sorted) {
                    $fullFile = Join-Path $sf.Directory $sf.FileName
                    $rel = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $fullFile
                    $targetsAll += [PSCustomObject]@{
                        FullPath      = $fullFile
                        RelPath       = $rel
                        PlatformLabel = $platformRoot.ToUpperInvariant()
                        DiskSort      = $sf.DiskSort
                        SideSort      = $sf.SideSort
                    }
                }

                # Ensures consistent <name> lines for NON-M3U disk entries when working with gamelist.xml
                Ensure-GameNamesInGamelist -State $state -Targets $targetsAll -FallbackName $playlistBase | Out-Null

                # If incomplete, unhide secondary entries and keep primary visible
                if ($setIsIncomplete) {

                    $toUnhideObjs = @($sorted | Where-Object { (Join-Path $_.Directory $_.FileName) -ne $primaryFull })
                    if ($toUnhideObjs.Count -gt 0) {

                        $targetsU = @()
                        foreach ($sf in $toUnhideObjs) {
                            $fullFile = Join-Path $sf.Directory $sf.FileName
                            $rel = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $fullFile
                            $targetsU += [PSCustomObject]@{
                                FullPath      = $fullFile
                                RelPath       = $rel
                                PlatformLabel = $platformRoot.ToUpperInvariant()
                                DiskSort      = $sf.DiskSort
                                SideSort      = $sf.SideSort
                            }
                        }

                        $unhideResult = Unhide-GameEntriesInGamelist -State $state -Targets $targetsU -UsedFiles $usedFiles

                        $platLabel = $platformRoot.ToUpperInvariant()
                        if (-not $gamelistUnhiddenCounts.ContainsKey($platLabel)) { $gamelistUnhiddenCounts[$platLabel] = 0 }
                        if (-not $gamelistAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistAlreadyVisibleCounts[$platLabel] = 0 }
                        $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
                        $gamelistAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount

                        foreach ($sf in $toUnhideObjs) {
                            $full = Join-Path $sf.Directory $sf.FileName
                            $usedFiles[$full] = $true
                        }
                    }

                    [void]$noM3UPrimaryEntriesIncomplete.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Primary entry kept visible (disk set incomplete)"
                    })

                    continue
                }

                # Identify Disk 2+ entries to hide (secondary entries)
                $toHide = @(
                    $sorted | Where-Object { (Join-Path $_.Directory $_.FileName) -ne $primaryFull }
                )

                # If nothing to hide, just report primary kept visible
                if ($toHide.Count -eq 0) {
                    [void]$noM3UPrimaryEntriesOk.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Primary entry kept visible"
                    })
                    continue
                }

                # Build hide target objects for Disk 2+
                $targets = @()
                foreach ($sf in $toHide) {

                    $fullFile = Join-Path $sf.Directory $sf.FileName
                    $rel = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $fullFile

                    $platformLabel = $platformRoot.ToUpperInvariant()
                    $targets += [PSCustomObject]@{
                        FullPath      = $fullFile
                        RelPath       = $rel
                        PlatformLabel = $platformLabel
                        DiskSort      = $sf.DiskSort
                        SideSort      = $sf.SideSort
                    }
                }

                # SAFETY GATE: never hide anything unless the primary exists in gamelist.xml AND
                # there is at least one secondary entry present in gamelist.xml to hide.
                $primaryPresent = $false
                $pathIndex = $null
                if ($state.Exists -and $null -ne $state.Lines) {
                    $pathIndex = Get-GamelistPathIndex -Lines $state.Lines
                } else {
                    $pathIndex = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                }

                if (-not [string]::IsNullOrWhiteSpace($primaryRel) -and $state.Exists -and $null -ne $state.Lines) {
                    $primaryPresent = $pathIndex.Contains($primaryRel)
                }

                $secondaryPresentCount = 0
                if ($state.Exists -and $null -ne $state.Lines) {
                    foreach ($t in $targets) {
                        if ([string]::IsNullOrWhiteSpace($t.RelPath)) { continue }
                        if ($pathIndex.Contains($t.RelPath)) {
                            $secondaryPresentCount++
                        }
                    }
                }

                if (-not $primaryPresent -or $secondaryPresentCount -lt 1) {

                    # If we can't prove a real multi-entry set exists in gamelist.xml, don't hide.
                    # Also make sure the would-be secondaries are unhidden (cleanup / safety).
                    $unhideResult = Unhide-GameEntriesInGamelist -State $state -Targets $targets -UsedFiles $usedFiles

                    $platLabel = $platformRoot.ToUpperInvariant()
                    if (-not $gamelistUnhiddenCounts.ContainsKey($platLabel)) { $gamelistUnhiddenCounts[$platLabel] = 0 }
                    if (-not $gamelistAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistAlreadyVisibleCounts[$platLabel] = 0 }
                    $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
                    $gamelistAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount

                    [void]$noM3UPrimaryEntriesIncomplete.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Skipping hide (primary missing from gamelist.xml or no secondary entries present in gamelist.xml)"
                    })

                    continue
                }

                # Track which entries SHOULD be hidden for this NON-M3U platform (Phase 3.5 reconciliation)
                if (-not $noM3UShouldBeHiddenByPlatform.ContainsKey($platformLower)) {
                    $noM3UShouldBeHiddenByPlatform[$platformLower] = @{}
                }
                foreach ($t in $targets) {
                    if (-not [string]::IsNullOrWhiteSpace($t.FullPath)) {
                        $noM3UShouldBeHiddenByPlatform[$platformLower][$t.FullPath] = $t.RelPath
                    }
                }

                # Apply hide operations in gamelist.xml
                $hideResult = Hide-GameEntriesInGamelist -State $state -Targets $targets -UsedFiles $usedFiles

                # Update per-platform hide counters
                $platLabel = $platformRoot.ToUpperInvariant()
                if (-not $gamelistHiddenCounts.ContainsKey($platLabel)) { $gamelistHiddenCounts[$platLabel] = 0 }
                if (-not $gamelistAlreadyHiddenCounts.ContainsKey($platLabel)) { $gamelistAlreadyHiddenCounts[$platLabel] = 0 }
                $gamelistHiddenCounts[$platLabel] += [int]$hideResult.NewlyHiddenCount
                $gamelistAlreadyHiddenCounts[$platLabel] += [int]$hideResult.AlreadyHiddenCount

                # Report primary visible outcome based on missing gamelist entries
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

                # Record per-platform count for NON-M3U sets
                $platformLabel = $platformRoot.ToUpperInvariant()
                if (-not $platformCounts.ContainsKey($platformLabel)) { $platformCounts[$platformLabel] = 0 }
                $platformCounts[$platformLabel]++
                $totalPlaylistsCreated++

                continue
            }

            # For M3U mode, skip incomplete sets
            if ($setIsIncomplete) { continue }

            # Write M3U file (unless dry run)
            if (-not $dryRun) {
                [System.IO.File]::WriteAllText($playlistPath, $cleanText, [System.Text.UTF8Encoding]::new($false))
            }

            $m3uWrittenPlaylistPaths[$playlistPath] = $true

            # Mark selected ROM files as used
            foreach ($sf in $sorted) {
                $full = Join-Path $sf.Directory $sf.FileName
                $usedFiles[$full] = $true
            }

            # Increment platform counters
            $platformLabel = Get-PlatformCountLabel -Directory $directory -ScriptDir $scriptDir
            if (-not $platformCounts.ContainsKey($platformLabel)) { $platformCounts[$platformLabel] = 0 }
            $platformCounts[$platformLabel]++
            $totalPlaylistsCreated++
        }
    }
}

# ==================================================================================================
# PHASE 3.5: RECONCILIATION PASS (NON-M3U GAMELIST VISIBILITY)
# ==================================================================================================

# Display reconciliation banner
Write-Phase "Reconciling NON-M3U gamelist visibility..."

# Only reconcile in XML mode
if ($noM3UPlatformMode -ieq "XML") {

    # Resolve NON-M3U platform set (lowercase)
    $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })

    foreach ($platformLower in @($noM3UPlatformsEncountered.Keys | Sort-Object -Unique)) {

        if (-not ($noM3USetLower -contains $platformLower.ToLowerInvariant())) { continue }

        $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
        $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

        # Build "should hide" set for this platform (FullPath -> RelPath)
        $shouldHideMap = @{}
        if ($noM3UShouldBeHiddenByPlatform.ContainsKey($platformLower)) {
            $shouldHideMap = $noM3UShouldBeHiddenByPlatform[$platformLower]
        }

        # Build candidate list: all parsed multi-disk candidates that reside in the PLATFORM ROOT (not subfolders)
        $candidates = @()
        foreach ($p in $parsed) {

            $plat = $null
            try {
                $plat = Get-PlatformRootName -Directory $p.Directory -ScriptDir $scriptDir
            } catch {
                $plat = $null
            }

            if ($null -eq $plat) { continue }
            if ($plat.ToLowerInvariant() -ne $platformLower.ToLowerInvariant()) { continue }

            # Only reconcile at the platform root directory (safety)
            try {
                $dirFull  = (Resolve-Path -LiteralPath $p.Directory).Path.TrimEnd('\')
                $rootFull = (Resolve-Path -LiteralPath $rootPath).Path.TrimEnd('\')
                if ($dirFull -ne $rootFull) { continue }
            } catch {
                continue
            }

            $full = Join-Path $p.Directory $p.FileName
            $rel  = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $full

            # Only reconcile entries we can map to a gamelist-relative path
            if ([string]::IsNullOrWhiteSpace($rel)) { continue }

            $candidates += [PSCustomObject]@{
                FullPath      = $full
                RelPath       = $rel
                PlatformLabel = $platformLower.ToUpperInvariant()
                DiskSort      = $p.DiskSort
                SideSort      = $p.SideSort
            }
        }

        if ($candidates.Count -eq 0) { continue }

        # Partition into "must be hidden" vs "must be visible"
        $toHide = @()
        $toUnhide = @()

        foreach ($c in $candidates) {

            $mustHide = $false
            if ($shouldHideMap.ContainsKey($c.FullPath)) { $mustHide = $true }

            if ($mustHide) { $toHide += $c }
            else { $toUnhide += $c }
        }

        # Apply unhide first (safety: never leave stale hidden behind)
        if ($toUnhide.Count -gt 0) {

            $unhideResult = Unhide-GameEntriesInGamelist -State $state -Targets $toUnhide -UsedFiles $usedFiles -SuppressAlreadyReport

            $platLabel = $platformLower.ToUpperInvariant()
            if (-not $gamelistUnhiddenCounts.ContainsKey($platLabel)) { $gamelistUnhiddenCounts[$platLabel] = 0 }
            if (-not $gamelistAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistAlreadyVisibleCounts[$platLabel] = 0 }
            $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
            $gamelistAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount
        }

        # Apply hide for entries that are still expected to be hidden
        if ($toHide.Count -gt 0) {

            $hideResult = Hide-GameEntriesInGamelist -State $state -Targets $toHide -UsedFiles $usedFiles -SuppressAlreadyReport

            $platLabel = $platformLower.ToUpperInvariant()
            if (-not $gamelistHiddenCounts.ContainsKey($platLabel)) { $gamelistHiddenCounts[$platLabel] = 0 }
            if (-not $gamelistAlreadyHiddenCounts.ContainsKey($platLabel)) { $gamelistAlreadyHiddenCounts[$platLabel] = 0 }
            $gamelistHiddenCounts[$platLabel] += [int]$hideResult.NewlyHiddenCount
            $gamelistAlreadyHiddenCounts[$platLabel] += [int]$hideResult.AlreadyHiddenCount
        }
    }
}

# ==================================================================================================
# PHASE 4: FLUSH GAMELIST CHANGES (NON-M3U PLATFORMS)
# ==================================================================================================

# Display finalization banner
Write-Phase "Finalizing gamelist.xml updates (if any)..."

# Save each cached gamelist if it was modified
foreach ($k in $gamelistStateByPlatform.Keys) {
    $st = $gamelistStateByPlatform[$k]
    if ($null -ne $st) {
        Save-GamelistIfChanged -State $st | Out-Null
    }
}

# ==================================================================================================
# PHASE 5: STRUCTURED REPORTING (SPLIT: M3U PLAYLISTS vs GAMELIST(S) UPDATED)
# ==================================================================================================

# Compute whether any playlist activity occurred
$anyM3UActivity =
    (@($m3uWrittenPlaylistPaths.Keys).Count -gt 0) -or
    (@($suppressedPreExistingPlaylists.Keys).Count -gt 0) -or
    (@($suppressedDuplicatePlaylists.Keys).Count -gt 0)

# Compute whether any gamelist activity occurred
$anyGamelistActivity =
    (@($noM3UPrimaryEntriesOk).Count -gt 0) -or
    (@($noM3UPrimaryEntriesIncomplete).Count -gt 0) -or
    (@($noM3UNoDisk1Sets).Count -gt 0) -or
    (@($noM3UNewlyHidden).Count -gt 0) -or
    (@($noM3UAlreadyHidden).Count -gt 0) -or
    (@($noM3UNewlyUnhidden).Count -gt 0) -or
    (@($noM3UAlreadyVisible).Count -gt 0) -or
    (@($noM3UMissingGamelistEntries).Count -gt 0) -or
    ($gamelistHiddenCounts.Count -gt 0) -or
    ($gamelistAlreadyHiddenCounts.Count -gt 0) -or
    ($gamelistUnhiddenCounts.Count -gt 0) -or
    ($gamelistAlreadyVisibleCounts.Count -gt 0)

# If no activity, print "nothing found" message
if (-not $anyM3UActivity -and -not $anyGamelistActivity) {
    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green
    Write-Host "No viable multi-disk files were found to create playlists from." -ForegroundColor Yellow
} else {

    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green

    # List created M3U playlists (and overwrite markers)
    if (@($m3uWrittenPlaylistPaths.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "CREATED" -ForegroundColor Green

        $m3uWrittenPlaylistPaths.Keys | Sort-Object | ForEach-Object {

            $p = $_

            if ($overwrittenExistingPlaylists.ContainsKey($p)) {
                Write-Host "$p" -NoNewline
                Write-Host " — Overwrote existing playlist that contained content discrepancy" -ForegroundColor Yellow
            } else {

                if ($dryRun) {
                    Write-Host "$p" -NoNewline
                    Write-Host " — DRY RUN (would write)" -ForegroundColor Yellow
                } else {
                    Write-Host $p
                }
            }
        }
    }
    else {
        if ($dryRun) {
            Write-Host "No M3U playlists were written (DRY RUN)." -ForegroundColor Yellow
        } else {
            Write-Host "No M3U playlists were written." -ForegroundColor Yellow
        }
    }

    # List suppressed pre-existing playlists
    if (@($suppressedPreExistingPlaylists.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "SUPPRESSED (PRE-EXISTING PLAYLIST CONTAINED IDENTICAL CONTENT)" -ForegroundColor Green
        $suppressedPreExistingPlaylists.Keys | Sort-Object | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
    }

    # List duplicate-collision suppressed playlists
    if (@($suppressedDuplicatePlaylists.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "SUPPRESSED (DUPLICATE CONTENT COLLISION DURING THIS RUN)" -ForegroundColor Green
        $suppressedDuplicatePlaylists.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-Host "$($_.Key)" -NoNewline -ForegroundColor Gray
            Write-Host " — Identical content collision with $($_.Value)" -ForegroundColor Yellow
        }
    }

    # If gamelist activity occurred, emit gamelist sections
    if ($anyGamelistActivity) {
        Write-Host ""

        if ($dryRun) {
            Write-Host "GAMELIST(S) UPDATED (DRY RUN — NO FILES MODIFIED)" -ForegroundColor Green
        } else {
            Write-Host "GAMELIST(S) UPDATED" -ForegroundColor Green
        }

        # Determine which gamelist.xml files were changed
        $changedGamelists = @()
        foreach ($k in $gamelistStateByPlatform.Keys) {
            $st = $gamelistStateByPlatform[$k]
            if ($null -ne $st -and $st.Exists -and $st.Changed -and $null -ne $st.GamelistPath) {
                $changedGamelists += $st.GamelistPath
            }
        }

        # Print changed gamelist paths
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
            Write-Host "No gamelist.xml files required modification." -ForegroundColor Yellow
        }

        # Print NON-M3U primary-visible OK bucket
        if (@($noM3UPrimaryEntriesOk).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (PRIMARY ENTRIES KEPT VISIBLE — OK)" -ForegroundColor Green
            $noM3UPrimaryEntriesOk | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print NON-M3U primary-visible incomplete bucket
        if (@($noM3UPrimaryEntriesIncomplete).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (PRIMARY ENTRIES KEPT VISIBLE — SET INCOMPLETE)" -ForegroundColor Green
            $noM3UPrimaryEntriesIncomplete | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print NON-M3U no-disk1 bucket
        if (@($noM3UNoDisk1Sets).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (NO DISK 1 — SET INCOMPLETE)" -ForegroundColor Green
            $noM3UNoDisk1Sets | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print newly hidden bucket
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

        # Print already hidden bucket
        if (@($noM3UAlreadyHidden).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES ALREADY HIDDEN (NO CHANGE)" -ForegroundColor Green
            $noM3UAlreadyHidden | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print newly unhidden bucket
        if (@($noM3UNewlyUnhidden).Count -gt 0) {
            Write-Host ""
            Write-Host "UNHIDDEN ENTRIES (NEW)" -ForegroundColor Green
            $noM3UNewlyUnhidden | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                if ($dryRun) {
                    Write-Host " — DRY RUN (would unhide in gamelist.xml)" -ForegroundColor Yellow
                } else {
                    Write-Host " — $($_.Reason)" -ForegroundColor Yellow
                }
            }
        }

        # Print already visible bucket
        if (@($noM3UAlreadyVisible).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES ALREADY VISIBLE (NO CHANGE)" -ForegroundColor Green
            $noM3UAlreadyVisible | Sort-Object FullPath | Group-Object FullPath | ForEach-Object {
                $x = $_.Group[0]
                Write-Host "$($x.FullPath)" -NoNewline
                Write-Host " — $($x.Reason)" -ForegroundColor Yellow
            }
        }

        # Print missing-from-gamelist bucket
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

# Initialize skipped-files report bucket
$notInPlaylists = [System.Collections.ArrayList]::new()

# Group parsed candidates by TitleKey to detect orphaned/unselected disk members
$groupsForNotUsed = $parsed | Group-Object TitleKey
foreach ($g in $groupsForNotUsed) {

    $gFiles = $g.Group

    # Handle singleton "Disk 2+" style or declared totals
    if (@($gFiles).Count -lt 2) {

        $f = $gFiles[0]
        $full = Join-Path $f.Directory $f.FileName

        $shouldReportSingleton =
            (($f.TotalDisks -ne $null -and $f.TotalDisks -ge 2) -or
             ($f.DiskSort -ne $null -and $f.DiskSort -ge 2))

        if ($shouldReportSingleton -and (-not $usedFiles.ContainsKey($full))) {

            if ($noM3UMissingGamelistEntryByFullPath.ContainsKey($full)) {
                [void]$notInPlaylists.Add([PSCustomObject]@{
                    FullPath = $full
                    Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
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

    # Compute observed disk set for this title
    $diskSet = @(
        $gFiles |
            Where-Object { $_.DiskSort -ne $null } |
            Select-Object -ExpandProperty DiskSort |
            Sort-Object -Unique
    )
    if ($diskSet.Count -eq 0) { continue }

    # Determine expected disks from declared total or max disk observed
    $maxTotal = ($gFiles |
        Where-Object { $_.TotalDisks -ne $null } |
        Select-Object -ExpandProperty TotalDisks |
        Sort-Object -Descending |
        Select-Object -First 1)

    $expectedDisks = @()
    if ($null -ne $maxTotal -and $maxTotal -ne "") { $expectedDisks = 1..([int]$maxTotal) }
    else {
        $maxDisk = ($diskSet | Sort-Object -Descending | Select-Object -First 1)
        if ($null -ne $maxDisk) { $expectedDisks = 1..([int]$maxDisk) } else { $expectedDisks = $diskSet }
    }

    # Build alt distribution lookup by disk number
    $altByDisk = @{}
    foreach ($x in $gFiles) {
        if ($null -eq $x.DiskSort) { continue }
        $dsk = [int]$x.DiskSort
        $a = Normalize-Alt $x.AltTag
        if (-not $altByDisk.ContainsKey($dsk)) { $altByDisk[$dsk] = @() }
        if (-not ($altByDisk[$dsk] -contains $a)) { $altByDisk[$dsk] += $a }
    }

    # Compute disk total mismatch bounds
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

    # Identify missing disks vs expectation
    $missingDisks = @()
    foreach ($ed in $expectedDisks) {
        if (-not ($diskSet -contains $ed)) { $missingDisks += $ed }
    }

    # For each unused file, compute skip reason
    foreach ($f in $gFiles) {

        $full = Join-Path $f.Directory $f.FileName
        if ($usedFiles.ContainsKey($full)) { continue }

        if ($noM3UMissingGamelistEntryByFullPath.ContainsKey($full)) {
            [void]$notInPlaylists.Add([PSCustomObject]@{
                FullPath = $full
                Reason   = "No entry in gamelist.xml (run gamelist update, scrape the game or manually add it into gamelist.xml)"
            })
            continue
        }

        # Ignore subfolder NON-M3U singles for skip reporting (safety)
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
                }
            }
        }

        $reason = "Unselected during fill"

        # If a different file for the same disk was selected, note it
        if ($f.DiskSort -ne $null) {

            $winner = $gFiles | Where-Object {
                $_.DiskSort -eq $f.DiskSort -and
                $usedFiles.ContainsKey((Join-Path $_.Directory $_.FileName))
            } | Select-Object -First 1

            if ($null -ne $winner) {
                $winnerFull = Join-Path $winner.Directory $winner.FileName
                $reason = "Unselected during fill (Disk $($f.DiskSort) chosen instead: $winnerFull)"
            }
        }

        # Apply [!] suppression reason label when it explains omission
        $hasBangInThisFile = ($f.BaseTagsKey -match '\[\!\]')
        $altNorm = Normalize-Alt $f.AltTag
        if (-not $hasBangInThisFile -and $altNorm) {

            $kBang = $f.TitleKey + "`0" + $f.BaseTagsKeyNB
            $kBangAlt = $f.TitleKey + "`0" + $f.BaseTagsKeyNB + "`0" + $altNorm

            if ($bangByTitleNB.ContainsKey($kBang) -and $bangAltByTitleNB.ContainsKey($kBangAlt)) {
                $reason = "Suppressed by [!] rule"
            }
        }

        # Prefer missing-disk label when disks are absent
        if ($reason -eq "Unselected during fill" -and $missingDisks.Count -gt 0) {
            $reason = "Missing matching disk"
        }

        # Prefer alt-fallback failure label when alt exists and no disks missing
        if ($reason -eq "Unselected during fill" -and $missingDisks.Count -eq 0 -and $altNorm) {
            $reason = "Alt fallback failed"
        }

        # Add mismatch detail when alt isn't present across disks
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

        # Collapse missing-matching-disk into incomplete-set reason
        if ($reason -eq "Missing matching disk" -and $missingDisks.Count -gt 0) {
            $reason = "Incomplete disk set"
        }

        # Specialize incomplete-set reason for missing Disk 1
        if ($reason -eq "Incomplete disk set" -and ($missingDisks -contains 1)) {
            $reason = "Incomplete disk set (missing Disk 1)"
        }

        # Add disk total mismatch note when totals disagree
        if ($reason -eq "Unselected during fill" -and
            $null -ne $minTotal -and $null -ne $maxTotalLocal -and
            $minTotal -ne $maxTotalLocal -and
            $f.DiskSort -gt [int]$minTotal) {

            $reason += " due to disk total mismatch"
        }

        [void]$notInPlaylists.Add([PSCustomObject]@{
            FullPath = $full
            Reason   = $reason
        })
    }
}

# Print skipped multi-disk files report (if any)
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

# Print playlist creation counts if any were recorded
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

# Print hidden-entry counts only when total > 0
$gamelistHiddenTotal = 0
foreach ($kv in $gamelistHiddenCounts.GetEnumerator()) { $gamelistHiddenTotal += [int]$kv.Value }
if ($gamelistHiddenTotal -gt 0) {
    Write-Host ""
    Write-Host "GAMELIST HIDDEN ENTRY COUNT(S)" -ForegroundColor Green
    $gamelistHiddenCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Name):" -ForegroundColor Cyan -NoNewline
        Write-Host " $($_.Value)"
    }
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $gamelistHiddenTotal"
}

# Print already-hidden counts only when total > 0
$gamelistAlreadyHiddenTotal = 0
foreach ($kv in $gamelistAlreadyHiddenCounts.GetEnumerator()) { $gamelistAlreadyHiddenTotal += [int]$kv.Value }
if ($gamelistAlreadyHiddenTotal -gt 0) {
    Write-Host ""
    Write-Host "GAMELIST ALREADY-HIDDEN COUNT(S)" -ForegroundColor Green
    $gamelistAlreadyHiddenCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Name):" -ForegroundColor Cyan -NoNewline
        Write-Host " $($_.Value)"
    }
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $gamelistAlreadyHiddenTotal"
}

# Print unhidden counts only when total > 0
$gamelistUnhiddenTotal = 0
foreach ($kv in $gamelistUnhiddenCounts.GetEnumerator()) { $gamelistUnhiddenTotal += [int]$kv.Value }
if ($gamelistUnhiddenTotal -gt 0) {
    Write-Host ""
    Write-Host "GAMELIST UNHIDDEN ENTRY COUNT(S)" -ForegroundColor Green
    $gamelistUnhiddenCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Name):" -ForegroundColor Cyan -NoNewline
        Write-Host " $($_.Value)"
    }
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $gamelistUnhiddenTotal"
}

# Print suppressed-playlist counts if any
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

# Print total skipped-file count if any
if (@($notInPlaylists).Count -gt 0) {
    Write-Host ""
    Write-Host "MULTI-DISK FILE SKIP COUNT" -ForegroundColor Green
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $(@($notInPlaylists).Count)"
}

# ==================================================================================================
# PHASE 8: FINAL RUNTIME REPORT
# ==================================================================================================

# Compute total elapsed runtime
$elapsed = (Get-Date) - $scriptStart
$totalSeconds = [int][math]::Floor($elapsed.TotalSeconds)

# Format runtime as seconds / M:SS / H:MM:SS
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

# Print runtime line
Write-Host ""
Write-Host "Runtime:" -ForegroundColor White -NoNewline
Write-Host " $runtimeText"
