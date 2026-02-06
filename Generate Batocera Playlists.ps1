<#
PURPOSE: Create .m3u files for each multi-disk game and insert the list of game filenames into the playlist or update gamelist.xml
VERSION: 1.8
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Requires PowerShell version 7+

- Place this file into the ROMS folder to process all platforms, or in a platform's individual subfolder to process just that one.

- Run the following in Batocera BEFORE running the script to ensure gamelist.xlm files are current and to avoid potential conflicts:
    - GAME SETTINGS > UPDATE GAMELISTS
    - SYSTEM SETTINGS > FRONTEND DEVELOPER OPTIONS > CLEAN GAMELISTS & REMOVE UNUSED MEDIA

- Run the following in Batocera AFTER running the script to ensure game lists are accurately displayed:
    - GAME SETTINGS > UPDATE GAMELISTS

OPTIONS:
- There are three variables you can update to change script behavior:
    - $dryRun : default is $false; set to $true to preview output without modifying files
    - $nonM3UPlatforms : array of platforms (folder names) that should NOT use M3U; default: 3DO + apple2
    - $noM3UPlatformMode : default "XML"; set to "skip" to ignore NON-M3U platforms entirely

BREAKDOWN:
- Enumerates ROM game files starting in the directory the script resides in
    - Scans up to 2 subdirectory levels deep recursively
    - Skips .m3u files during scanning (so it doesn’t treat playlists as input)
    - Skips common media/manual folders (e.g., images, videos, media, manuals, downloaded_*) to reduce false multi-disk detections

- Detects multi-disk candidates by parsing filenames for “designators”
    - A designator is a disk/disc/side marker that indicates a set (case-insensitive), such as:
        - Disk 1, Disc B, Disk II, Disk 2 of 6, Disk 4 Side A, etc.
        - Side-only sets like Side A, Side B, etc. are supported (treated as Disk 1 with different sides)
    - Supports disk tokens as:
        - Numbers (1, 2, …)
        - Letters (A, B, …)
        - Roman numerals (I … XX) (used for sort normalization)
    - Also recognizes optional patterns like:
        - “of N” totals (e.g., Disk 2 of 6)
        - Side X paired with a disk marker (e.g., Disk 2 Side B)

- Normalizes and groups detected files into candidate multi-disk sets
    - Determines logical DiskSort values for ordering and comparison
    - Ensures Disk 1 is present for a set to be considered complete
    - Treats sets missing Disk 1 as incomplete and suppresses playlist creation

- Resolves ambiguous multi-disk sets with mixed file variants
    - Prevents multiple files with the same disk number (e.g., two Disc 1s) from entering the same set
    - When multiple extensions exist for the same title (e.g., .cdi and .chd):
        - Selects a single “winning” extension set
        - Preference is given to the extension set spanning the greatest number of distinct disks
    - Non-selected variants are treated as incomplete and excluded from playlist creation

- Creates .m3u playlist files for valid multi-disk sets
    - Orders playlist entries by normalized disk order
    - Writes relative paths suitable for Batocera
    - Suppresses playlist creation when an existing .m3u already contains identical content
    - Tracks suppressed playlists separately from newly created ones
    - Supports execution from either a platform folder or the ROMS root

- Updates gamelist.xml entries for M3U playlists
    - Creates a new <game> entry when a playlist exists but no gamelist entry is present
    - Marks script-managed playlist entries with a <dtw_m3u_entry>true</dtw_m3u_entry> tag
    - Clones metadata from Disk 1 when creating a new playlist entry
    - Ensures newly created entries are reported as “filled” during the same run

- Reconciles existing M3U gamelist entries
    - Fills missing metadata/media from Disk 1 only when:
        - The destination tag is missing or invalid, AND
        - The source value differs from what already exists
    - Skips filling when existing values are already identical to the source
    - Prevents repeated refilling of unchanged metadata on subsequent runs
    - Ensures identical behavior whether run from ROMS or a single platform folder

- Handles playlist-owned media assets
    - Copies Disk 1 media files (image, video, marquee, thumbnail, bezel) so the playlist owns its own media
    - Validates that source media files exist before copying
    - Does not invent media references when source files do not exist
    - Retargets playlist metadata to the copied media paths when successful
    - Leaves media tags blank when no valid source media exists

- Preserves non-media metadata correctly
    - Copies descriptive tags (name, desc, developer, publisher, genre, etc.) from Disk 1
    - Does not apply filesystem existence checks to non-file metadata
    - Treats Disk 1’s <name> as authoritative when propagating titles

- For platforms that can't use M3U playlist files, the script instead hides Disk 2+ in gamelist.xml (<hidden>true</hidden>)
    - Per run, a backup of gamelist.xml is made first called gamelist.backup
    - If a platform is added into $nonM3UPlatforms, the script will also delete any existing .m3u files under that platform (respecting $dryRun)
    - If a platform is removed from $nonM3UPlatforms (becoming M3U again), the script will unhide only the Disk 2+ entries it previously hid (tracked via a custom marker tag)
    - This creates a single entry in Batocera for a multi-disk game instead of one per disk
    - Initially includes 3DO and Apple II, but additional platform folders can be added into $nonM3UPlatforms
    - Ensures the canonical <name> exists across Disk 2+ entries
        - If Disk 1 already has a <name>, it is propagated to all discs in the set
    - If you’d rather skip these platforms entirely, set $noM3UPlatformMode to "skip" instead of "XML"

- Supports dry-run mode
    - Performs full detection and reporting without writing files or modifying gamelist.xml

- Generates detailed end-of-run reporting
    - Reports:
        - Newly created playlists
        - Suppressed playlists
        - M3U gamelist entries created or filled
        - Gamelist entries already visible
        - Multi-disk files skipped due to ambiguity or incompleteness
    - Counts are reported per platform and as totals, using consistent formatting
#>

# ==================================================================================================
# PHASE 0: SCRIPT STARTUP / CONFIG / STATE
# ==================================================================================================

# Establish script working directory and start time
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# --------------------------------------------------------------------------------------------------
# Ensure all relative paths resolve from the script's folder (critical when running from ROMS root).
# Also restore the original location even if the script errors out.
# --------------------------------------------------------------------------------------------------
$__dtw_originalLocation = Get-Location
$__dtw_locationPushed = $false

trap {
    # Preserve original error and ensure we restore the user's working directory.
    $err = $_

    try {
        if ($__dtw_locationPushed) { Pop-Location }
    } catch {
        # ignore cleanup errors
    }

    try {
        if ($null -ne $__dtw_originalLocation -and -not [string]::IsNullOrWhiteSpace($__dtw_originalLocation.Path)) {
            Set-Location -LiteralPath $__dtw_originalLocation.Path
        }
    } catch {
        # ignore cleanup errors
    }

    throw $err
}

Push-Location -LiteralPath $scriptDir
$__dtw_locationPushed = $true

$scriptStart = Get-Date

# Initialize platform playlist count buckets and total counter
$platformCounts = @{}
$totalPlaylistsCreated = 0

# User config: dry run toggle
$dryRun = $false # <-- set to $true if you want to see what the output would be without changing files

# User config: platforms that should not use M3U playlists
$nonM3UPlatforms = @(
    '3DO'
    'apple2'
)

# User config: NON-M3U platform handling mode
$noM3UPlatformMode = "XML"   # <-- set to "skip" to completely ignore those platforms

# Attempt to widen console buffer (best-effort)
try {
    $raw = $Host.UI.RawUI
    $size = $raw.BufferSize
    # If buffer width is narrow, expand it
    if ($size.Width -lt 300) {
        $raw.BufferSize = New-Object Management.Automation.Host.Size(300, $size.Height)
    }
} catch {
    # Ignore console resize failures
}

# Declare folder names that should not be scanned
$skipFolders = @(
    'images','videos','media','manuals',
    'downloaded_images','downloaded_videos','downloaded_media','downloaded_manuals'
)

# Initialize cached gamelist state per platform
$gamelistStateByPlatform = @{}     # platformLower -> state object (cached lines)

# Initialize per-run gamelist backup tracking
$gamelistBackupDone      = @{}     # gamelist path -> $true

# Initialize a lookup of "missing from gamelist.xml" by full ROM path
$noM3UMissingGamelistEntryByFullPath = @{}  # full file path -> $true

# Initialize NON-M3U reporting buckets (ArrayList for safety)
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

# Initialize per-platform count buckets for gamelist changes
$gamelistHiddenCounts          = @{}  # platform label -> count newly hidden
$gamelistAlreadyHiddenCounts   = @{}  # platform label -> count already hidden (no change)

$gamelistUnhiddenCounts        = @{}  # platform label -> count newly unhidden
$gamelistEntriesAlreadyVisibleCounts  = @{}  # platform label -> count already visible (no change)
$totalGamelistEntriesAlreadyVisible   = 0

# Track NON-M3U entries that SHOULD be hidden per platform
$noM3UShouldBeHiddenByPlatform = @{}  # platformLower -> hashtable fullPath -> relPath
$noM3UPlatformsEncountered     = @{}  # platformLower -> $true

# Marker tag used to track which hidden entries were hidden by THIS script (for safe reclassification unhide)
$dtwNonM3UMarkerTagName = "dtw_nonm3u_hidden"
$dtwNonM3UMarkerLine    = "<$dtwNonM3UMarkerTagName>true</$dtwNonM3UMarkerTagName>"

# Marker tag used to track which M3U playlist entries were created/managed by THIS script
$dtwM3UMarkerTagName = "dtw_m3u_entry"
$dtwM3UMarkerLine    = "<$dtwM3UMarkerTagName>true</$dtwM3UMarkerTagName>"

# Collect M3U entries that can be repaired from Disk 1 (fill-only)
$m3uRepairQueue = [System.Collections.ArrayList]::new() # PSCustomObject { PlatformLower; PlatformRootPath; PlaylistPath; PlaylistRel; Disk1Rel; PlaylistBaseName }

# Reporting: marker tag operations (retro-tag + cleanup/unhide)
$noM3UMarkerNewlyAdded   = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UMarkerRemoved      = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UMarkerAddedSet     = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$noM3UMarkerRemovedSet   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

# Reporting: M3U deletion when platform is NON-M3U
$nonM3UDeletedM3UFiles    = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$nonM3UDeletedM3UCounts   = @{}  # platform label -> count deleted

# Reporting: M3U playlist entries filled/repaired in gamelist.xml
$m3uEntriesFilled         = [System.Collections.ArrayList]::new() # PSCustomObject { Platform; PlaylistRel; Reason }
$m3uEntriesFilledSet      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

# --------------------------------------------------------------------------------------------------
# M3U MEDIA REUSE SUPPORT (OPTION A)
# - When an M3U playlist is created, clone Disk 1’s gamelist.xml entry (if present) into a new entry
#   for the .m3u, and copy/rename referenced media files (image/video/etc.) so the M3U entry owns them.
# - This keeps Batocera “Clean gamelists & remove unused media” safe, since the M3U entry references them.
# --------------------------------------------------------------------------------------------------

# Common media tags we will clone/copy for M3U playlist entries (batocera/emulationstation variants)
$m3uMediaTagNames = @(
    'image','thumbnail','video','marquee','fanart','boxart','cartridge','title','mix','bezel'
)

# Common non-media tags we will fill from Disk 1 into an M3U entry (fill-only)
$m3uMetadataTagNames = @(
    'desc','genre','releasedate','developer','publisher','players','rating','lang','region'
)

# ==================================================================================================
# FUNCTIONS
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
        $parts = @([string]$rel -split '\\')
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
    if (-not [string]::IsNullOrWhiteSpace($rel)) { $parts = @([string]$rel -split '\\') }

    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # If ROMs root, label as PLATFORM\subpath
    if ($scriptIsRomsRoot) {
        if (@($parts).Count -eq 0) { return $scriptLeaf.ToUpperInvariant() }
        $platform = @($parts)[0].ToUpperInvariant()
        $subParts = if (@($parts).Count -gt 1) { @(@($parts)[1..(@($parts).Count-1)]) } else { @() }
        if (@($subParts).Count -gt 0) { return ($platform + "\" + ($subParts -join "\")) }
        return $platform
    }

    # If platform root, label as PLATFORM\subpath
    $platform = $scriptLeaf.ToUpperInvariant()
    if (@($parts).Count -gt 0) { return ($platform + "\" + ($parts -join "\")) }
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
        # Root must exist; file may not (e.g., M3U path during creation)
        $rootFull = (Resolve-Path -LiteralPath $PlatformRootPath).Path
        $rootFull = [System.IO.Path]::GetFullPath($rootFull).TrimEnd('\','/')

        # Do NOT Resolve-Path the file (it may not exist yet)
        $fileFull = [System.IO.Path]::GetFullPath($FileFullPath)
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

# Generate a gamelist.backup path (always overwrites)
function Get-UniqueGamelistBackupPath {
    param([Parameter(Mandatory=$true)][string]$GamelistPath)

    $dir = Split-Path -Parent $GamelistPath
    $base = Join-Path $dir "gamelist.backup"

    # Always overwrite the same backup file
    return $base
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

# Delete all .m3u files under a NON-M3U platform root (respecting $dryRun)
function Delete-M3UFilesUnderPlatformRoot {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformLabel,
        [Parameter(Mandatory=$true)][string]$PlatformRootPath
    )

    if (-not (Test-Path -LiteralPath $PlatformRootPath)) { return 0 }

    $deletedCount = 0

    # Match script scan depth and skip folders
    $m3us = @()
    try {
        $m3us = @(Get-ChildItem -LiteralPath $PlatformRootPath -File -Recurse -Depth 2 -Filter "*.m3u" -ErrorAction SilentlyContinue)
    } catch {
        $m3us = @()
    }

    foreach ($f in $m3us) {

        # Safety: don't delete anything inside known media/manual folders
        if (Is-InSkipFolderPath -FileFullPath $f.FullName -SkipFolderNames $skipFolders) { continue }

        if ($dryRun) {
            [void]$nonM3UDeletedM3UFiles.Add([PSCustomObject]@{
                FullPath = $f.FullName
                Reason   = "DRY RUN (would delete .m3u because platform is NON-M3U)"
            })
            $deletedCount++
            continue
        }

        try {
            Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
            [void]$nonM3UDeletedM3UFiles.Add([PSCustomObject]@{
                FullPath = $f.FullName
                Reason   = "Deleted .m3u because platform is NON-M3U"
            })
            $deletedCount++
        } catch {
            # If delete fails, still report it (but do not count as deleted)
            [void]$nonM3UDeletedM3UFiles.Add([PSCustomObject]@{
                FullPath = $f.FullName
                Reason   = "FAILED to delete .m3u (platform is NON-M3U): $($_.Exception.Message)"
            })
        }
    }

    if (-not $nonM3UDeletedM3UCounts.ContainsKey($PlatformLabel)) { $nonM3UDeletedM3UCounts[$PlatformLabel] = 0 }
    $nonM3UDeletedM3UCounts[$PlatformLabel] += [int]$deletedCount

    return $deletedCount
}

# Unhide ONLY entries that are marked as hidden by this script (marker tag) for a platform now treated as M3U
function Unhide-MarkedEntriesInPlatformGamelist {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)][string]$PlatformLabel,
        [Parameter(Mandatory=$true)][hashtable]$UsedFiles
    )

    $result = [PSCustomObject]@{
        DidWork            = $false
        NewlyUnhiddenCount = 0
        MarkerRemovedCount = 0
    }

    if (-not $State.Exists -or $null -eq $State.Lines) { return $result }

    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $State.Lines) { [void]$lines.Add($ln) }

    for ($i = 0; $i -lt $lines.Count; $i++) {

        $line = $lines[$i]
        if ($null -eq $line) { continue }

        $m = [regex]::Match($line, '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
        if (-not $m.Success) { continue }

        $indent = $m.Groups['I'].Value
        $rel = $m.Groups['V'].Value

        # Walk this <game> block
        $j = $i + 1
        $endIndex = $null
        $hiddenIndex = $null
        $hiddenValue = $null
        $markerIndex = $null

        while ($j -lt $lines.Count) {

            $tline = $lines[$j]
            if ($tline -match '^\s*<path>\s*') { break }
            if ($tline -match '^\s*</game>\s*$') { $endIndex = $j; break }

            if ($null -eq $hiddenIndex) {
                $hm = [regex]::Match($tline, '^\s*<hidden>\s*(?<H>.*?)\s*</hidden>\s*$')
                if ($hm.Success) { $hiddenIndex = $j; $hiddenValue = $hm.Groups['H'].Value }
            }

            if ($null -eq $markerIndex) {
                $mm = [regex]::Match($tline, '^\s*<' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*(?<M>.*?)\s*</' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*$')
                if ($mm.Success) { $markerIndex = $j }
            }

            $j++
        }

        # Only act if marker exists
        if ($null -eq $markerIndex) { continue }

        # Remove marker line always (cleanup)
        if (-not $dryRun) {
            $lines.RemoveAt($markerIndex)
            $State.Changed = $true
        }
        $result.DidWork = $true
        $result.MarkerRemovedCount++

        # Report marker removal (micro-dedupe by rel path where possible)
        if ($noM3UMarkerRemovedSet.Add([string]$rel)) {
            [void]$noM3UMarkerRemoved.Add([PSCustomObject]@{
                FullPath = $rel
                Reason   = "Removed NON-M3U marker because platform is now treated as M3U"
            })
        }

        # If marker was before hidden line, hidden index shifts left by 1 after removal
        if ($null -ne $hiddenIndex -and $markerIndex -lt $hiddenIndex) { $hiddenIndex-- }

        # If hidden=true exists, remove it (unhide)
        if ($null -ne $hiddenIndex -and $hiddenValue -match '^(?i)true$') {

            if (-not $dryRun) {
                $lines.RemoveAt($hiddenIndex)
                $State.Changed = $true
            }
            $result.DidWork = $true
            $result.NewlyUnhiddenCount++

            # We only have a rel path here; still mark "used" defensively as rel path
            $UsedFiles[$rel] = $true

            [void]$noM3UNewlyUnhidden.Add([PSCustomObject]@{
                FullPath = $rel
                Reason   = "Unhidden because platform is now treated as M3U (marker-based)"
            })
        }

        # Continue scan safely: rewind slightly (block mutations)
        if ($i -gt 0) { $i-- }
    }

    $State.Lines = $lines.ToArray()
    return $result
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

    # ---- Use List[string] for efficient inserts/edits ----
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
        # ---- Explicitly treat missing name as mismatch ----
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

    # ---- Use List[string] for efficient inserts/edits ----
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

            # Walk within this <game> block to find existing <hidden> and DTW marker
            $j = $i + 1
            $hiddenLineIndex = $null
            $hiddenValue = $null
            $markerLineIndex = $null

            while ($j -lt $lines.Count) {
                $tline = $lines[$j]

                if ($tline -match '^\s*<path>\s*') { break }
                if ($tline -match '^\s*</game>\s*$') { break }

                if ($null -eq $hiddenLineIndex) {
                    $hm = [regex]::Match($tline, '^\s*<hidden>\s*(?<H>.*?)\s*</hidden>\s*$')
                    if ($hm.Success) {
                        $hiddenLineIndex = $j
                        $hiddenValue = $hm.Groups['H'].Value
                    }
                }

                if ($null -eq $markerLineIndex) {
                    $mm = [regex]::Match(
                        $tline,
                        '^\s*<' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*(?<M>.*?)\s*</' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*$'
                    )
                    if ($mm.Success) { $markerLineIndex = $j }
                }

                if ($null -ne $hiddenLineIndex -and $null -ne $markerLineIndex) { break }

                $j++
            }

            # If <hidden> exists, ensure it is true; else insert a new line
            if ($null -ne $hiddenLineIndex) {

                if ($hiddenValue -match '^(?i)true$') {
                    $UsedFiles[$t.FullPath] = $true

                    # Retro-tag: if already hidden but marker is missing, add marker so reclassification can safely unhide later
                    if ($null -eq $markerLineIndex) {
                        if (-not $dryRun) {
                            # Insert marker immediately after <hidden> line
                            $lines.Insert($hiddenLineIndex + 1, ($indent + $dtwNonM3UMarkerLine))
                            $State.Changed = $true
                        }

                        $result.DidWork = $true

                        if ($noM3UMarkerAddedSet.Add([string]$t.FullPath)) {
                            [void]$noM3UMarkerNewlyAdded.Add([PSCustomObject]@{
                                FullPath = $t.FullPath
                                Reason   = "Added NON-M3U marker to an already-hidden entry (retro-tag for safe reclassification)"
                            })
                        }
                    }

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
                            $State.Changed = $true
                        }
                        $result.DidWork = $true
                    }

                    $handled = $true
                }
                else {
                    if (-not $dryRun) {
                        $lines[$hiddenLineIndex] = ($lines[$hiddenLineIndex] -replace '(?i)<hidden>\s*.*?\s*</hidden>', '<hidden>true</hidden>')
                        $State.Changed = $true
                    }

                    # Ensure marker exists when we hide
                    if ($null -eq $markerLineIndex) {
                        if (-not $dryRun) {
                            $lines.Insert($hiddenLineIndex + 1, ($indent + $dtwNonM3UMarkerLine))
                            $State.Changed = $true
                        }

                        if ($noM3UMarkerAddedSet.Add([string]$t.FullPath)) {
                            [void]$noM3UMarkerNewlyAdded.Add([PSCustomObject]@{
                                FullPath = $t.FullPath
                                Reason   = "Added NON-M3U marker (entry hidden by this script)"
                            })
                        }
                    }

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
                            $State.Changed = $true
                        }
                        $result.DidWork = $true
                    }

                    $handled = $true
                }
            }
            else {

                if (-not $dryRun) {
                    $insertLine = ($indent + "<hidden>true</hidden>")
                    $lines.Insert($i + 1, $insertLine)

                    # Insert marker immediately after <hidden>
                    $lines.Insert($i + 2, ($indent + $dtwNonM3UMarkerLine))

                    $State.Changed = $true
                }

                if ($noM3UMarkerAddedSet.Add([string]$t.FullPath)) {
                    [void]$noM3UMarkerNewlyAdded.Add([PSCustomObject]@{
                        FullPath = $t.FullPath
                        Reason   = "Added NON-M3U marker (entry hidden by this script)"
                    })
                }

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
                        $State.Changed = $true
                    }
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

    # ---- Use List[string] for efficient deletes/edits ----
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
                    $State.Changed = $true
                }
                $result.DidWork = $true
            }

            # Walk within this <game> block to find existing <hidden> and DTW marker
            $j = $i + 1
            $hiddenLineIndex = $null
            $hiddenValue = $null
            $markerLineIndex = $null

            while ($j -lt $lines.Count) {
                $tline = $lines[$j]

                if ($tline -match '^\s*<path>\s*') { break }
                if ($tline -match '^\s*</game>\s*$') { break }

                if ($null -eq $hiddenLineIndex) {
                    $hm = [regex]::Match($tline, '^\s*<hidden>\s*(?<H>.*?)\s*</hidden>\s*$')
                    if ($hm.Success) {
                        $hiddenLineIndex = $j
                        $hiddenValue = $hm.Groups['H'].Value
                    }
                }

                if ($null -eq $markerLineIndex) {
                    $mm = [regex]::Match(
                        $tline,
                        '^\s*<' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*(?<M>.*?)\s*</' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*$'
                    )
                    if ($mm.Success) { $markerLineIndex = $j }
                }

                if ($null -ne $hiddenLineIndex -and $null -ne $markerLineIndex) { break }

                $j++
            }

            # If hidden=true, remove that line; else record already visible
            if ($null -ne $hiddenLineIndex -and $hiddenValue -match '^(?i)true$') {

                # Remove <hidden>true</hidden>
                if (-not $dryRun) {
                    $lines.RemoveAt($hiddenLineIndex)
                    $State.Changed = $true
                }

                # Also remove marker if present (cleanup)
                if ($null -ne $markerLineIndex) {

                    # If marker was after hidden, its index shifts by -1 after removing hidden
                    if ($markerLineIndex -gt $hiddenLineIndex) { $markerLineIndex-- }

                    if (-not $dryRun) {
                        $lines.RemoveAt($markerLineIndex)
                        $State.Changed = $true
                    }

                    if ($noM3UMarkerRemovedSet.Add([string]$t.FullPath)) {
                        [void]$noM3UMarkerRemoved.Add([PSCustomObject]@{
                            FullPath = $t.FullPath
                            Reason   = "Removed NON-M3U marker (entry unhidden)"
                        })
                    }
                }

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

                # Cleanup: if marker exists but entry is not hidden, remove marker
                if ($null -ne $markerLineIndex) {

                    if (-not $dryRun) {
                        $lines.RemoveAt($markerLineIndex)
                        $State.Changed = $true
                    }

                    $result.DidWork = $true

                    if ($noM3UMarkerRemovedSet.Add([string]$t.FullPath)) {
                        [void]$noM3UMarkerRemoved.Add([PSCustomObject]@{
                            FullPath = $t.FullPath
                            Reason   = "Removed NON-M3U marker from an already-visible entry (cleanup)"
                        })
                    }
                }

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

# Extract a <game> block (as lines) for a given <path> rel value
function Get-GameBlockByRelPath {
    param(
        [Parameter(Mandatory=$true)]$Lines,
        [Parameter(Mandatory=$true)][string]$RelPath
    )

    if ($null -eq $Lines -or [string]::IsNullOrWhiteSpace($RelPath)) { return $null }

    # Find the <path> line
    for ($i = 0; $i -lt $Lines.Count; $i++) {

        $line = $Lines[$i]
        if ($null -eq $line) { continue }

        $m = [regex]::Match($line, '^\s*<path>\s*(?<V>.*?)\s*</path>\s*$')
        if (-not $m.Success) { continue }
        if ($m.Groups['V'].Value -ine $RelPath) { continue }

        # Find <game ...> start by scanning backwards (safe cap)
        $start = $null
        $backCap = 80
        $bMin = [Math]::Max(0, $i - $backCap)

        for ($b = $i; $b -ge $bMin; $b--) {
            if ($Lines[$b] -match '^\s*<game\b') { $start = $b; break }
            if ($Lines[$b] -match '^\s*</game>\s*$') { break }
        }

        if ($null -eq $start) { return $null }

        # Find </game> end by scanning forward (safe cap)
        $end = $null
        $fCap = 400
        $fMax = [Math]::Min($Lines.Count - 1, $start + $fCap)

        for ($f = $start; $f -le $fMax; $f++) {
            if ($Lines[$f] -match '^\s*</game>\s*$') { $end = $f; break }
        }

        if ($null -eq $end) { return $null }

        # Return block + indices
        $block = @()
        for ($k = $start; $k -le $end; $k++) { $block += $Lines[$k] }

        return [PSCustomObject]@{
            StartIndex = $start
            EndIndex   = $end
            BlockLines = $block
        }
    }

    return $null
}

# Find an existing <game> block for a given rel path; returns start/end indices or $null
function Find-GameBlockRangeByRelPath {
    param(
        [Parameter(Mandatory=$true)]$Lines,
        [Parameter(Mandatory=$true)][string]$RelPath
    )

    $blk = Get-GameBlockByRelPath -Lines $Lines -RelPath $RelPath
    if ($null -eq $blk) { return $null }

    return [PSCustomObject]@{
        StartIndex = $blk.StartIndex
        EndIndex   = $blk.EndIndex
    }
}

# Copy a media file referenced in gamelist.xml and return the new rel path (or original if copy not possible)
function Copy-M3UMediaFileAndReturnRelPath {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformRootPath,
        [Parameter(Mandatory=$true)][string]$SourceRelPath,
        [Parameter(Mandatory=$true)][string]$PlaylistBaseName,
        [Parameter(Mandatory=$false)][string]$TagName
    )

    if ([string]::IsNullOrWhiteSpace($PlatformRootPath)) { return $null }
    if ([string]::IsNullOrWhiteSpace($SourceRelPath))   { return $null }
    if ([string]::IsNullOrWhiteSpace($PlaylistBaseName)) { return $null }

    # Only handle relative-style paths (./something)
    if (-not ($SourceRelPath -match '^\./')) { return $null }

    # Convert to filesystem path
    $srcRelNoDot = $SourceRelPath.Substring(2) # remove "./"
    $srcRelWin   = ($srcRelNoDot -replace '/', '\')
    $srcFull     = Join-Path $PlatformRootPath $srcRelWin

    if (-not (Test-Path -LiteralPath $srcFull)) { return $null }

    $srcDir = Split-Path -Parent $srcFull
    $ext    = [System.IO.Path]::GetExtension($srcFull)
    if ([string]::IsNullOrWhiteSpace($ext)) { return $null }

    # Determine Batocera suffix by tag (if known)
    $suffix = $null
    switch -Regex ($TagName) {
        '^image$'     { $suffix = '-image'; break }
        '^thumbnail$' { $suffix = '-thumb'; break }
        '^marquee$'   { $suffix = '-marquee'; break }
        '^bezel$'     { $suffix = '-bezel'; break }
        '^video$'     { $suffix = '-video'; break }
        default       { $suffix = $null; break }
    }

    # Dest filename:
    # - if tag suffix is known, use PlaylistBaseName + suffix + ext
    # - else fall back to old behavior (PlaylistBaseName + ext)
    $destName = $PlaylistBaseName + $ext
    if (-not [string]::IsNullOrWhiteSpace($suffix)) {
        $destName = $PlaylistBaseName + $suffix + $ext
    }

    # Dest: same folder as source, but named to match the playlist (so the M3U “owns” its media)
    $destFull = Join-Path $srcDir $destName

    # If already exists, keep it
    if (Test-Path -LiteralPath $destFull) {
        $base = (Resolve-Path -LiteralPath $PlatformRootPath).Path.TrimEnd('\')
        $rel = $destFull.Substring($base.Length).TrimStart('\')
        $rel = $rel -replace '\\', '/'
        return ("./" + $rel)
    }

    # Respect dry run
    if ($dryRun) { return $SourceRelPath }

    try {
        Copy-Item -LiteralPath $srcFull -Destination $destFull -Force
        $base = (Resolve-Path -LiteralPath $PlatformRootPath).Path.TrimEnd('\')
        $rel = $destFull.Substring($base.Length).TrimStart('\')
        $rel = $rel -replace '\\', '/'
        return ("./" + $rel)
    } catch {
        return $null
    }
}

# Clone Disk 1 gamelist entry into a playlist entry, copying media so the playlist owns it.
function Ensure-M3UPlaylistEntryWithMedia {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)][string]$PlatformRootPath,
        [Parameter(Mandatory=$true)][string]$Disk1RelPath,
        [Parameter(Mandatory=$true)][string]$PlaylistRelPath,
        [Parameter(Mandatory=$true)][string]$PlaylistBaseName
    )

    if (-not $State.Exists -or $null -eq $State.Lines) { return $false }

    # If the playlist entry already exists, do nothing (minimal / safe)
    $already = Find-GameBlockRangeByRelPath -Lines $State.Lines -RelPath $PlaylistRelPath
    if ($null -ne $already) { return $false }

    # Find Disk 1 entry block
    $src = Get-GameBlockByRelPath -Lines $State.Lines -RelPath $Disk1RelPath
    if ($null -eq $src -or $null -eq $src.BlockLines -or $src.BlockLines.Count -lt 2) { return $false }

    # Clone lines
    $newBlock = @()
    foreach ($ln in $src.BlockLines) { $newBlock += $ln }

    # Replace <path> to playlist rel
    for ($i = 0; $i -lt $newBlock.Count; $i++) {
        $m = [regex]::Match($newBlock[$i], '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
        if ($m.Success) {
            $indent = $m.Groups['I'].Value
            $newBlock[$i] = ($indent + "<path>$PlaylistRelPath</path>")
            break
        }
    }

    # Ensure <name> exists (prefer Disk 1 <name>, else PlaylistBaseName)
    $disk1Name = $null
    
    # Prefer extracting <name> directly from the Disk 1 block we are cloning from (more reliable during initial creation)
    $srcBlock = $src.BlockLines
    if ($srcBlock -and $srcBlock.Count -gt 0) {
        foreach ($ln in $srcBlock) {
            if ($ln -match '^\s*<name>(.*?)</name>\s*$') {
                $disk1Name = $matches[1]
                break
            }
        }
    }

    # Fallback to lookup-by-path if needed
    if ([string]::IsNullOrWhiteSpace($disk1Name)) {
        $disk1Name = Get-GamelistNameByRelPath -Lines $State.Lines -RelPath $Disk1RelPath
    }

    $useName = $disk1Name
    if ([string]::IsNullOrWhiteSpace($useName)) { $useName = $PlaylistBaseName }

    $nameFound = $false
    for ($i = 0; $i -lt $newBlock.Count; $i++) {
        $nm = [regex]::Match($newBlock[$i], '^(?<I>\s*)<name>\s*(?<N>.*?)\s*</name>\s*$')
        if ($nm.Success) {
            $indent = $nm.Groups['I'].Value
            $newBlock[$i] = ($indent + "<name>$useName</name>")
            $nameFound = $true
            break
        }
    }

    if (-not $nameFound) {
        # Insert name immediately after <path>
        for ($i = 0; $i -lt $newBlock.Count; $i++) {
            if ($newBlock[$i] -match '^\s*<path>\s*') {
                $indent = ([regex]::Match($newBlock[$i], '^(?<I>\s*)').Groups['I'].Value)
                $newBlock = @($newBlock[0..$i] + @($indent + "<name>$useName</name>") + $newBlock[($i+1)..($newBlock.Count-1)])
                break
            }
        }
    }

    # Remove <hidden>true</hidden> and NON-M3U marker if present (playlist entry should be visible)
    $filtered = @()
    foreach ($ln in $newBlock) {
        if ($ln -match '^\s*<hidden>\s*(?i:true)\s*</hidden>\s*$') { continue }
        if ($ln -match '^\s*<' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*.*?\s*</' + [regex]::Escape($dtwNonM3UMarkerTagName) + '>\s*$') { continue }
        $filtered += $ln
    }
    $newBlock = $filtered

    # Ensure M3U marker tag exists in the cloned playlist block
    $hasM3UMarker = $false
    foreach ($ln in $newBlock) {
        if ($ln -match '^\s*<' + [regex]::Escape($dtwM3UMarkerTagName) + '>\s*(?i:true)\s*</' + [regex]::Escape($dtwM3UMarkerTagName) + '>\s*$') {
            $hasM3UMarker = $true
            break
        }
    }

    if (-not $hasM3UMarker) {
        for ($i = 0; $i -lt $newBlock.Count; $i++) {
            if ($newBlock[$i] -match '^\s*<path>\s*') {
                $indent = ([regex]::Match($newBlock[$i], '^(?<I>\s*)').Groups['I'].Value)
                $newBlock = @($newBlock[0..$i] + @($indent + $dtwM3UMarkerLine) + $newBlock[($i+1)..($newBlock.Count-1)])
                break
            }
        }
    }

    # Copy/retarget media tags so playlist owns its media
    for ($i = 0; $i -lt $newBlock.Count; $i++) {

        $line = $newBlock[$i]
        if ($null -eq $line) { continue }

        foreach ($tag in $m3uMediaTagNames) {

            $rx = '^(?<I>\s*)<' + [regex]::Escape($tag) + '>\s*(?<V>.*?)\s*</' + [regex]::Escape($tag) + '>\s*$'
            $mm = [regex]::Match($line, $rx)
            if (-not $mm.Success) { continue }

            $indent = $mm.Groups['I'].Value
            $val    = $mm.Groups['V'].Value

            # Copy and get new rel path
            $newRel = Copy-M3UMediaFileAndReturnRelPath -PlatformRootPath $PlatformRootPath -SourceRelPath $val -PlaylistBaseName $PlaylistBaseName -TagName $tag
            $newBlock[$i] = ($indent + "<$tag>$newRel</$tag>")
            break
        }
    }

    # Insert new block just before </gameList> if present, else append at end
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $State.Lines) { [void]$lines.Add($ln) }

    $insertAt = $lines.Count
    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        if ($lines[$i] -match '^\s*</gameList>\s*$') { $insertAt = $i; break }
    }

    if (-not $dryRun) {
        for ($k = 0; $k -lt $newBlock.Count; $k++) {
            $lines.Insert($insertAt + $k, $newBlock[$k])
        }
        $State.Lines = $lines.ToArray()
        $State.Changed = $true
    }



    # Record reporting entry (micro-dedupe per playlist) so first-run creation is reported (and not only subsequent repairs)
    $reportKey = ($PlatformRootPath + "`0" + $PlaylistRelPath)
    if ($m3uEntriesFilledSet.Add($reportKey)) {
        [void]$m3uEntriesFilled.Add([PSCustomObject]@{
            Platform    = (Split-Path -Leaf $PlatformRootPath)
            PlaylistRel = $PlaylistRelPath
            Reason      = "Created playlist entry from Disk 1 (clone)"
        })
    }
    return $true
}

# Fill missing M3U playlist entry fields from Disk 1 (fill-only; no overwrites)
function Repair-M3UPlaylistEntryFromDisk1 {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)][string]$PlatformRootPath,
        [Parameter(Mandatory=$true)][string]$Disk1RelPath,
        [Parameter(Mandatory=$true)][string]$PlaylistRelPath,
        [Parameter(Mandatory=$true)][string]$PlaylistBaseName,
        [Parameter(Mandatory=$true)][string]$PlaylistPath
    )

    if (-not $State.Exists -or $null -eq $State.Lines) { return $false }

    # Safety: only repair when the playlist file exists on disk
    if ([string]::IsNullOrWhiteSpace($PlaylistPath) -or -not (Test-Path -LiteralPath $PlaylistPath)) { return $false }

    # Locate playlist block and disk 1 block
    $pl = Get-GameBlockByRelPath -Lines $State.Lines -RelPath $PlaylistRelPath
    if ($null -eq $pl -or $null -eq $pl.BlockLines -or $pl.BlockLines.Count -lt 2) { return $false }

    $src = Get-GameBlockByRelPath -Lines $State.Lines -RelPath $Disk1RelPath
    if ($null -eq $src -or $null -eq $src.BlockLines -or $src.BlockLines.Count -lt 2) { return $false }

    # Build src tag lookup (simple tag -> value)
    $srcTag = @{}
    foreach ($ln in $src.BlockLines) {
        if ($null -eq $ln) { continue }
        $m = [regex]::Match($ln, '^\s*<(?<T>[A-Za-z0-9_]+)>\s*(?<V>.*?)\s*</\k<T>>\s*$')
        if ($m.Success) {
            $t = $m.Groups['T'].Value
            $v = $m.Groups['V'].Value
            if (-not $srcTag.ContainsKey($t)) { $srcTag[$t] = $v }
        }
    }

    # Work on a mutable copy of playlist block lines
    $blk = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $pl.BlockLines) { [void]$blk.Add($ln) }

    # Determine indent for inserts (prefer <path> indent)
    $insertIndent = ""
    foreach ($ln in $blk) {
        $pm = [regex]::Match($ln, '^(?<I>\s*)<path>\s*')
        if ($pm.Success) { $insertIndent = $pm.Groups['I'].Value; break }
    }

    # Detect existing marker (script-managed)
    $hasMarker = $false
    for ($i = 0; $i -lt $blk.Count; $i++) {
        if ($blk[$i] -match '^\s*<' + [regex]::Escape($dtwM3UMarkerTagName) + '>\s*(?i:true)\s*</' + [regex]::Escape($dtwM3UMarkerTagName) + '>\s*$') {
            $hasMarker = $true
            break
        }
    }

    # Helper: find tag line index in a block
    function Find-TagLineIndex {
        param([System.Collections.Generic.List[string]]$B, [string]$TagName)
        for ($x = 0; $x -lt $B.Count; $x++) {
            $rx = '^\s*<' + [regex]::Escape($TagName) + '>\s*(?<V>.*?)\s*</' + [regex]::Escape($TagName) + '>\s*$'
            if ([regex]::Match($B[$x], $rx).Success) { return $x }
        }
        return $null
    }

    # Helper: get tag value if present
    function Get-TagValue {
        param([System.Collections.Generic.List[string]]$B, [string]$TagName)
        $ix = Find-TagLineIndex -B $B -TagName $TagName
        if ($null -eq $ix) { return $null }
        $m = [regex]::Match($B[$ix], '^\s*<' + [regex]::Escape($TagName) + '>\s*(?<V>.*?)\s*</' + [regex]::Escape($TagName) + '>\s*$')
        if ($m.Success) { return $m.Groups['V'].Value }
        return $null
    }

    # Helper: insert a tag line after <path> (or after marker if present), if missing
    function Insert-TagAfterPath {
        param([System.Collections.Generic.List[string]]$B, [string]$Indent, [string]$TagName, [string]$Value)

        $at = $null
        for ($x = 0; $x -lt $B.Count; $x++) {
            if ($B[$x] -match '^\s*<path>\s*') { $at = $x; break }
        }
        if ($null -eq $at) { return $false }

        # If marker exists directly after <path>, insert after it
        $after = $at + 1
        if ($after -lt $B.Count -and $B[$after] -match '^\s*<' + [regex]::Escape($dtwM3UMarkerTagName) + '>\s*') {
            $after++
        }

        $B.Insert($after, ($Indent + "<$TagName>$Value</$TagName>"))
        return $true
    }

    $didWork = $false

    # Ensure marker exists (retro-tag allowed only when mapping is known via the repair queue)
    if (-not $hasMarker) {
        if (-not $dryRun) {
            # Insert marker immediately after <path>
            for ($x = 0; $x -lt $blk.Count; $x++) {
                if ($blk[$x] -match '^\s*<path>\s*') {
                    $ind = ([regex]::Match($blk[$x], '^(?<I>\s*)').Groups['I'].Value)
                    $blk.Insert($x + 1, ($ind + $dtwM3UMarkerLine))
                    $didWork = $true
                    break
                }
            }
        } else {
            $didWork = $true
        }
        $hasMarker = $true
    }

    \
    # Always set <name> from Disk 1 (but only mark as changed when different)
    $srcName = $null
    if ($srcTag.ContainsKey("name")) { $srcName = $srcTag["name"] }
    $useName = $srcName
    if ([string]::IsNullOrWhiteSpace($useName)) { $useName = $PlaylistBaseName }

    if (-not [string]::IsNullOrWhiteSpace($useName)) {
        $curName = Get-TagValue -B $blk -TagName "name"
        if ($null -eq $curName -or $curName -ne $useName) {
            $ix = Find-TagLineIndex -B $blk -TagName "name"
            if ($null -ne $ix) {
                $ind = ([regex]::Match($blk[$ix], '^(?<I>\s*)').Groups['I'].Value)
                if (-not $dryRun) { $blk[$ix] = ($ind + "<name>$useName</name>") }
                $didWork = $true
            } else {
                if (-not $dryRun) { [void](Insert-TagAfterPath -B $blk -Indent $insertIndent -TagName "name" -Value $useName) }
                $didWork = $true
            }
        }
    }

    # Fill non-media metadata tags (only if missing or different)
 (only if missing or different)
    foreach ($tag in $m3uMetadataTagNames) {

        if (-not $srcTag.ContainsKey($tag)) { continue }
        $srcVal = $srcTag[$tag]
        if ([string]::IsNullOrWhiteSpace($srcVal)) { continue }

        $curVal = Get-TagValue -B $blk -TagName $tag

        # Skip if already identical
        if ($null -ne $curVal -and $curVal -eq $srcVal) {
            continue
        }

        $ix = Find-TagLineIndex -B $blk -TagName $tag
        if ($null -ne $ix) {
            $ind = ([regex]::Match($blk[$ix], '^(?<I>\s*)').Groups['I'].Value)
            if (-not $dryRun) {
                $blk[$ix] = ($ind + "<$tag>$srcVal</$tag>")
                }
            $didWork = $true
        } else {
            if (-not $dryRun) {
                [void](Insert-TagAfterPath -B $blk -Indent $insertIndent -TagName $tag -Value $srcVal)
            }
            $didWork = $true
        }
    }

    \
    # Fill media tags
    foreach ($tag in $m3uMediaTagNames) {

        if (-not $srcTag.ContainsKey($tag)) { continue }
        $srcVal = $srcTag[$tag]
        if ([string]::IsNullOrWhiteSpace($srcVal)) { continue }

        $curVal = Get-TagValue -B $blk -TagName $tag

        # If current value exists and points to a real file, keep it (even if different)
        $curExists = $false
        if (-not [string]::IsNullOrWhiteSpace($curVal)) {
            $curRel = $curVal.Trim()
            $curFsRel = $curRel
            if ($curFsRel.StartsWith('./')) { $curFsRel = $curFsRel.Substring(2) }
            if ($curFsRel.StartsWith('.\')) { $curFsRel = $curFsRel.Substring(2) }
            $curFsRel = $curFsRel -replace '/', '\'
            $curFull = Join-Path $PlatformRootPath $curFsRel
            if (Test-Path -LiteralPath $curFull) { $curExists = $true }
        }

        if ($curExists) { continue }

        # Compute the new rel path by copying/retargeting from Disk 1.
        $newRel = Copy-M3UMediaFileAndReturnRelPath -PlatformRootPath $PlatformRootPath -SourceRelPath $srcVal -PlaylistBaseName $PlaylistBaseName -TagName $tag
        if ([string]::IsNullOrWhiteSpace($newRel)) {
            # If current value is invalid and we can't produce a valid replacement, remove the tag to avoid stale references.
            $ixRemove = Find-TagLineIndex -B $blk -TagName $tag
            if ($null -ne $ixRemove) {
                if (-not $dryRun) { $blk.RemoveAt([int]$ixRemove) }
                $didWork = $true
            }
            continue
        }

        # Skip if already identical
        if ($null -ne $curVal -and $curVal -eq $newRel) { continue }

        $ix = Find-TagLineIndex -B $blk -TagName $tag
        if ($null -ne $ix) {
            $ind = ([regex]::Match($blk[$ix], '^(?<I>\s*)').Groups['I'].Value)
            if (-not $dryRun) { $blk[$ix] = ($ind + "<$tag>$newRel</$tag>") }
            $didWork = $true
        } else {
            if (-not $dryRun) { [void](Insert-TagAfterPath -B $blk -Indent $insertIndent -TagName $tag -Value $newRel) }
            $didWork = $true
        }
    }


    # If nothing changed, stop
    if (-not $didWork) { return $false }

    # Replace the block back into State.Lines
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $State.Lines) { [void]$lines.Add($ln) }

    if (-not $dryRun) {
        $removeCount = ($pl.EndIndex - $pl.StartIndex + 1)
        $lines.RemoveRange($pl.StartIndex, $removeCount)
        for ($k = 0; $k -lt $blk.Count; $k++) {
            $lines.Insert($pl.StartIndex + $k, $blk[$k])
        }
        $State.Lines = $lines.ToArray()
        $State.Changed = $true
    } else {
        $State.Changed = $true
    }

    # Record reporting entry ONLY if something actually changed
    if ($didWork) {
        $reportKey = ($PlatformRootPath + "`0" + $PlaylistRelPath)
        if ($m3uEntriesFilledSet.Add($reportKey)) {
            [void]$m3uEntriesFilled.Add([PSCustomObject]@{
                Platform    = (Split-Path -Leaf $PlatformRootPath)
                PlaylistRel = $PlaylistRelPath
                Reason      = "Filled missing metadata/media from Disk 1 (fill-only)"
            })
        }
    }

    return $true
}

""
Write-Host "If you have a lot of ROMs some phases may take a while to complete" -ForegroundColor DarkYellow

# ==================================================================================================
# PHASE 1: FILE ENUMERATION / PARSING
# ==================================================================================================

# Initialize parsed candidate collection
$parsed = @()

# Display scan banner
Write-Phase "Collecting ROM file data (scanning folders)..."

# If in XML mode, delete existing .m3u files for platforms currently designated NON-M3U (respect $dryRun)
if ($noM3UPlatformMode -ieq "XML") {

    $scriptFull = (Resolve-Path -LiteralPath $scriptDir).Path.TrimEnd('\')
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # Determine which NON-M3U platforms are in-scope for this run
    $scopeNoM3U = @()
    if ($scriptIsRomsRoot) {
        $scopeNoM3U = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
    } else {
        $scopeNoM3U = @($scriptLeaf.ToLowerInvariant())
    }

    foreach ($p in @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })) {

        if (-not ($scopeNoM3U -contains $p)) { continue }

        $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $p
        $platLabel = $p.ToUpperInvariant()

        Delete-M3UFilesUnderPlatformRoot -PlatformLabel $platLabel -PlatformRootPath $rootPath | Out-Null
    }
}

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

            # ------------------------------------------------------------------------------------------
            # - Prevent multiple Disc/Disk 1 variants (e.g., .cdi + .chd) from entering the same playlist.
            #   Pick a single “winning” extension set (most distinct disk numbers, prefers a set that has Disk 1).
            # ------------------------------------------------------------------------------------------
            $dupDiskNums = @($playlistFiles | Group-Object DiskSort | Where-Object { $_.Count -gt 1 })
            if ($dupDiskNums.Count -gt 0) {

                $byExt = @($playlistFiles | Group-Object { [System.IO.Path]::GetExtension($_.FileName).ToLowerInvariant() })

                if ($byExt.Count -gt 1) {

                    $best = $byExt | Sort-Object `
                        @{ Expression = { (@($_.Group | Where-Object { $_.DiskSort -eq 1 }).Count -gt 0) }; Descending = $true }, `
                        @{ Expression = { @($_.Group | Select-Object -ExpandProperty DiskSort -Unique).Count }; Descending = $true }, `
                        @{ Expression = { @($_.Group).Count }; Descending = $true } | Select-Object -First 1

                    if ($null -ne $best -and -not [string]::IsNullOrWhiteSpace($best.Name)) {

                        $keepExt = $best.Name
                        $filtered = @($playlistFiles | Where-Object { [System.IO.Path]::GetExtension($_.FileName).ToLowerInvariant() -eq $keepExt })

                        if ($filtered.Count -gt 0) {
                            $playlistFiles = $filtered
                        }
                    }
                }
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

            # Track encountered NON-M3U platforms
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
                    if (-not $gamelistEntriesAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistEntriesAlreadyVisibleCounts[$platLabel] = 0 }
                    $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
                    $gamelistEntriesAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount

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

            # Determine completeness (declared total OR inferred gap check)
            $setIsIncomplete = $false

            $presentDisks = @(
                $playlistFiles |
                    Where-Object { $_.DiskSort -ne $null } |
                    Select-Object -ExpandProperty DiskSort |
                    Sort-Object -Unique
            )

            $missingDisksLocal = @()

            if ($rootTotal) {

                foreach ($ed in (1..$rootTotal)) {
                    if (-not ($presentDisks -contains $ed)) { $missingDisksLocal += $ed }
                }

            } else {

                # Infer expectation from max observed disk to detect gaps (e.g., Disk 1 + Disk 3 => missing Disk 2)
                if ($presentDisks.Count -gt 0) {
                    $maxDiskLocal = ($presentDisks | Sort-Object -Descending | Select-Object -First 1)
                    if ($null -ne $maxDiskLocal -and [int]$maxDiskLocal -ge 2) {
                        foreach ($ed in (1..([int]$maxDiskLocal))) {
                            if (-not ($presentDisks -contains $ed)) { $missingDisksLocal += $ed }
                        }
                    }
                }
            }

            if ($missingDisksLocal.Count -gt 0) { $setIsIncomplete = $true }

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

            # Queue M3U repair mapping from Disk 1 (fill-only pass uses this mapping)
            if (-not $isNoM3U) {

                try {
                    $platformRootQ = Get-PlatformRootName -Directory $directory -ScriptDir $scriptDir
                    if ($null -ne $platformRootQ) {

                        $platformLowerQ = $platformRootQ.ToLowerInvariant()
                        $rootPathQ = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLowerQ

                        $disk1ObjQ = @($sorted | Where-Object { $_.DiskSort -eq 1 } | Sort-Object SideSort | Select-Object -First 1)
                        if ($null -ne $disk1ObjQ) {

                            $disk1FullQ = Join-Path $disk1ObjQ.Directory $disk1ObjQ.FileName
                            $disk1RelQ  = Get-RelativeGamelistPath -PlatformRootPath $rootPathQ -FileFullPath $disk1FullQ
                            $m3uRelQ    = Get-RelativeGamelistPath -PlatformRootPath $rootPathQ -FileFullPath $playlistPath

                            if (-not [string]::IsNullOrWhiteSpace($disk1RelQ) -and -not [string]::IsNullOrWhiteSpace($m3uRelQ)) {

                                # Avoid duplicate queue items within the same run
                                $dupKey = ($platformLowerQ + "`0" + $m3uRelQ)
                                $alreadyQueued = $false
                                foreach ($it in $m3uRepairQueue) {
                                    if (($it.PlatformLower + "`0" + $it.PlaylistRel) -ieq $dupKey) { $alreadyQueued = $true; break }
                                }

                                if (-not $alreadyQueued) {
                                    [void]$m3uRepairQueue.Add([PSCustomObject]@{
                                        PlatformLower    = $platformLowerQ
                                        PlatformRootPath = $rootPathQ
                                        PlaylistPath     = $playlistPath
                                        PlaylistRel      = $m3uRelQ
                                        Disk1Rel         = $disk1RelQ
                                        PlaylistBaseName = $playlistBase
                                    })
                                }
                            }
                        }
                    }
                } catch {
                    # Ignore mapping failures
                }
            }

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
                if (-not $gamelistEntriesAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistEntriesAlreadyVisibleCounts[$platLabel] = 0 }
                $gamelistUnhiddenCounts[$platLabel] += [int]$unhidePrimaryResult.NewlyUnhiddenCount
                $gamelistEntriesAlreadyVisibleCounts[$platLabel] += [int]$unhidePrimaryResult.AlreadyVisibleCount

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
                        if (-not $gamelistEntriesAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistEntriesAlreadyVisibleCounts[$platLabel] = 0 }
                        $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
                        $gamelistEntriesAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount

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
                    if (-not $gamelistEntriesAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistEntriesAlreadyVisibleCounts[$platLabel] = 0 }
                    $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
                    $gamelistEntriesAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount

                    [void]$noM3UPrimaryEntriesIncomplete.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Skipping hide (primary missing from gamelist.xml or no secondary entries present in gamelist.xml)"
                    })

                    continue
                }

                # Track which entries SHOULD be hidden for this NON-M3U platform
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
                if (-not $gamelistEntriesAlreadyVisibleCounts.ContainsKey($platformLabel)) { $gamelistEntriesAlreadyVisibleCounts[$platformLabel] = 0 }
$gamelistEntriesAlreadyVisibleCounts[$platformLabel]++
$totalGamelistEntriesAlreadyVisible++
continue
            }

            # For M3U mode, skip incomplete sets
            if ($setIsIncomplete) { continue }

            # Write M3U file (unless dry run)
            if (-not $dryRun) {
                [System.IO.File]::WriteAllText($playlistPath, $cleanText, [System.Text.UTF8Encoding]::new($false))
            }

            $m3uWrittenPlaylistPaths[$playlistPath] = $true

            # --------------------------------------------------------------------------------------------------
            # - If Disk 1 has a scraped gamelist.xml entry, clone it to a new entry for the .m3u
            #   and copy/rename media so the playlist entry "owns" its media.
            # --------------------------------------------------------------------------------------------------

            try {

                $platformRoot = Get-PlatformRootName -Directory $directory -ScriptDir $scriptDir
                if ($null -ne $platformRoot) {

                    $platformLower = $platformRoot.ToLowerInvariant()
                    $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower

                    $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

                    # Need Disk 1 rel path
                    $disk1Obj = @($sorted | Where-Object { $_.DiskSort -eq 1 } | Sort-Object SideSort | Select-Object -First 1)

                    if ($null -eq $disk1Obj) {
                        Write-Warn "Could not locate Disk 1 candidate for '$playlistBase' in '$platformLower'; skipping gamelist M3U entry creation."
                    }
                    else {
                        $disk1Full = Join-Path $disk1Obj.Directory $disk1Obj.FileName
                        $disk1Rel  = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $disk1Full
                        $m3uRel    = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $playlistPath

                        if ([string]::IsNullOrWhiteSpace($disk1Rel) -or [string]::IsNullOrWhiteSpace($m3uRel)) {
                            Write-Warn "Could not compute Disk1RelPath/PlaylistRelPath for '$playlistBase' in '$platformLower'; skipping gamelist M3U entry creation."
                        }
                        else {
                            Ensure-M3UPlaylistEntryWithMedia -State $state -PlatformRootPath $rootPath -Disk1RelPath $disk1Rel -PlaylistRelPath $m3uRel -PlaylistBaseName $playlistBase | Out-Null
                        }
                    }
                }

            } catch {
                # Intentionally swallow any media-reuse issues to avoid breaking playlist generation.
            }

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
# RECONCILIATION PASS (M3U GAMELIST METADATA)
# ==================================================================================================

Write-Phase "Reconciling M3U gamelist metadata..."

if (@($m3uRepairQueue).Count -gt 0) {

    foreach ($it in $m3uRepairQueue) {

        try {

            $platformLower = $it.PlatformLower
            $rootPath      = $it.PlatformRootPath
            $playlistRel   = $it.PlaylistRel
            $disk1Rel      = $it.Disk1Rel
            $playlistBase  = $it.PlaylistBaseName
            $playlistPath  = $it.PlaylistPath

            if ([string]::IsNullOrWhiteSpace($platformLower)) { continue }
            if ([string]::IsNullOrWhiteSpace($rootPath)) { continue }
            if ([string]::IsNullOrWhiteSpace($playlistRel)) { continue }
            if ([string]::IsNullOrWhiteSpace($disk1Rel)) { continue }
            if ([string]::IsNullOrWhiteSpace($playlistBase)) { continue }

            $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath
            if (-not $state.Exists -or $null -eq $state.Lines) { continue }

            # Only repair when the playlist entry already exists in gamelist.xml
            $existing = Find-GameBlockRangeByRelPath -Lines $state.Lines -RelPath $playlistRel
            if ($null -eq $existing) { continue }

            Repair-M3UPlaylistEntryFromDisk1 -State $state -PlatformRootPath $rootPath -Disk1RelPath $disk1Rel -PlaylistRelPath $playlistRel -PlaylistBaseName $playlistBase -PlaylistPath $playlistPath | Out-Null

        } catch {
            # Ignore repair failures
        }
    }
}

# ==================================================================================================
# RECONCILIATION PASS (NON-M3U GAMELIST VISIBILITY)
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
            if (-not $gamelistEntriesAlreadyVisibleCounts.ContainsKey($platLabel)) { $gamelistEntriesAlreadyVisibleCounts[$platLabel] = 0 }
            $gamelistUnhiddenCounts[$platLabel] += [int]$unhideResult.NewlyUnhiddenCount
            $gamelistEntriesAlreadyVisibleCounts[$platLabel] += [int]$unhideResult.AlreadyVisibleCount
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
# RECLASSIFICATION SWEEP (M3U PLATFORMS: UNHIDE ONLY MARKED ENTRIES)
# ==================================================================================================

Write-Phase "Reconciling M3U reclassification (unhiding only entries previously hidden by this script)..."

if ($noM3UPlatformMode -ieq "XML") {

    # Resolve NON-M3U platform set (lowercase)
    $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })

    $scriptFull = (Resolve-Path -LiteralPath $scriptDir).Path.TrimEnd('\')
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    # Determine platforms in scope
    $platformsToCheck = @()
    if ($scriptIsRomsRoot) {
        # Check every platform folder that actually exists under ROMs
        try {
            $platformsToCheck = @(Get-ChildItem -LiteralPath $scriptDir -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.Name.ToLowerInvariant() })
        } catch {
            $platformsToCheck = @()
        }
    } else {
        $platformsToCheck = @($scriptLeaf.ToLowerInvariant())
    }

    foreach ($platformLower in @($platformsToCheck | Sort-Object -Unique)) {

        # Only sweep platforms that are CURRENTLY treated as M3U (i.e., not in NON-M3U list)
        if ($noM3USetLower -contains $platformLower) { continue }

        $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
        $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

        if (-not $state.Exists -or $null -eq $state.Lines) { continue }

        $platLabel = $platformLower.ToUpperInvariant()
        $sweep = Unhide-MarkedEntriesInPlatformGamelist -State $state -PlatformLabel $platLabel -UsedFiles $usedFiles

        # Count these under the existing unhidden buckets (they will show as rel paths)
        if (-not $gamelistUnhiddenCounts.ContainsKey($platLabel)) { $gamelistUnhiddenCounts[$platLabel] = 0 }
        $gamelistUnhiddenCounts[$platLabel] += [int]$sweep.NewlyUnhiddenCount

        # Marker removals are reported via $noM3UMarkerRemoved
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
    ($gamelistEntriesAlreadyVisibleCounts.Count -gt 0)

# If no activity, print "nothing found" message
if (-not $anyM3UActivity -and -not $anyGamelistActivity) {
    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green
    Write-Host "No viable multi-disk files were found to create playlists from." -ForegroundColor Yellow
} else {

    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green

    # Report deleted M3U playlists for platforms configured as NON-M3U
    if ($null -ne $nonM3UDeletedM3UFiles -and @($nonM3UDeletedM3UFiles).Count -gt 0) {

        Write-Host ""
        Write-Host "DELETED (PLATFORM CONFIGURED AS NON-M3U)" -ForegroundColor Green

        $nonM3UDeletedM3UFiles | Sort-Object FullPath | ForEach-Object {
            Write-Host "$($_.FullPath)" -NoNewline
            Write-Host " — $($_.Reason)" -ForegroundColor Yellow
        }
    }

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
        Write-Host "PRE-EXISTING PLAYLIST CONTAINED IDENTICAL CONTENT (NO CHANGE)" -ForegroundColor Green
        $suppressedPreExistingPlaylists.Keys | Sort-Object | ForEach-Object { Write-Host $_ -ForegroundColor DarkYellow }
    }

    # List duplicate-collision suppressed playlists
    if (@($suppressedDuplicatePlaylists.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "DUPLICATE CONTENT COLLISION DURING THIS RUN (SUPPRESSED)" -ForegroundColor Green
        $suppressedDuplicatePlaylists.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-Host "$($_.Key)" -NoNewline -ForegroundColor DarkYellow
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

        # Print filled M3U entry bucket
        if (@($m3uEntriesFilled).Count -gt 0) {
            Write-Host ""
            Write-Host "M3U ENTRIES FILLED IN GAMELIST.XML (NEW)" -ForegroundColor Green
            $m3uEntriesFilled | Sort-Object Platform, PlaylistRel | ForEach-Object {
                Write-Host ("{0}\{1}" -f $_.Platform, $_.PlaylistRel) -NoNewline
                if ($dryRun) {
                    Write-Host " — DRY RUN (would fill)" -ForegroundColor Yellow
                } else {
                    Write-Host (" — {0}" -f $_.Reason) -ForegroundColor Yellow
                }
            }
        }

        # Print NON-M3U primary-visible OK bucket
        if (@($noM3UPrimaryEntriesOk).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U COMPLETE SETS IDENTIFIED (DISK 1 KEPT VISIBLE)" -ForegroundColor Green
            $noM3UPrimaryEntriesOk | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print NON-M3U primary-visible incomplete bucket
        if (@($noM3UPrimaryEntriesIncomplete).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U INCOMPLETE SETS IDENTIFIED (DISK 1 KEPT VISIBLE)" -ForegroundColor Green
            $noM3UPrimaryEntriesIncomplete | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print NON-M3U no-disk1 bucket
        if (@($noM3UNoDisk1Sets).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U INCOMPLETE SETS IDENTIFIED (NO DISK 1 FOUND)" -ForegroundColor Green
            $noM3UNoDisk1Sets | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print newly hidden bucket
        if (@($noM3UNewlyHidden).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES HIDDEN (NEW)" -ForegroundColor Green
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
                Write-Host "$($_.FullPath)" -NoNewline -ForegroundColor DarkYellow
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        # Print newly unhidden bucket
        if (@($noM3UNewlyUnhidden).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES UNHIDDEN (NEW)" -ForegroundColor Green
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
                Write-Host "$($x.FullPath)" -NoNewline  -ForegroundColor DarkYellow
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
    Write-Host "POSSIBLE MULTI-DISK FILES SKIPPED (ADDRESS MANUALLY)" -ForegroundColor Green
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
    # Recompute total from per-platform counts to avoid stale totals
    $totalPlaylistsCreated = 0
    foreach ($kvp in $platformCounts.GetEnumerator()) { $totalPlaylistsCreated += [int]$kvp.Value }

    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $totalPlaylistsCreated"
}

# Print NON-M3U visibility counts (platforms that cannot use M3U)
if ($gamelistEntriesAlreadyVisibleCounts.Count -gt 0 -or $totalGamelistEntriesAlreadyVisible -gt 0) {
    Write-Host ""
    Write-Host "GAMELIST.XML ENTRIES ALREADY VISIBLE" -ForegroundColor Green
    $gamelistEntriesAlreadyVisibleCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Name):" -ForegroundColor Cyan -NoNewline
        Write-Host " $($_.Value)"
    }
    # Recompute total from per-platform counts to avoid stale totals
    $totalGamelistEntriesAlreadyVisible = 0
    foreach ($kvp in $gamelistEntriesAlreadyVisibleCounts.GetEnumerator()) { $totalGamelistEntriesAlreadyVisible += [int]$kvp.Value }

    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $totalGamelistEntriesAlreadyVisible"
}


# Print filled M3U entry count
if (@($m3uEntriesFilled).Count -gt 0) {
    Write-Host ""
    Write-Host "M3U PLAYLIST ENTRIES FILLED COUNT(S)" -ForegroundColor Green

    $filledByPlatform = @($m3uEntriesFilled | Group-Object Platform | Sort-Object Name)
    $filledTotal = 0
    foreach ($grp in $filledByPlatform) {
        if ($null -eq $grp) { continue }
        $p = [string]$grp.Name
        if ([string]::IsNullOrWhiteSpace($p)) { $p = "UNKNOWN" }
        Write-Host ($p.ToUpperInvariant() + ":") -ForegroundColor Cyan -NoNewline
        Write-Host (" " + [string]$grp.Count) -ForegroundColor White
        $filledTotal += [int]$grp.Count
    }
    Write-Host ("TOTAL: {0}" -f $filledTotal)
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
    # Only show platforms with non-zero counts (avoid printing long lists of 0s)
    $gamelistUnhiddenCounts.GetEnumerator() | Where-Object { [int]$_.Value -gt 0 } | Sort-Object Name | ForEach-Object {
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
Write-Host "Runtime:" -ForegroundColor DarkYellow -NoNewline
Write-Host " $runtimeText" -ForegroundColor DarkYellow

# Restore original working directory (see trap at top)
if ($__dtw_locationPushed) { Pop-Location }
Set-Location -LiteralPath $__dtw_originalLocation.Path
