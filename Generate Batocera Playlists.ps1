<#
PURPOSE: Create .m3u files for each multi-disk game and insert the list of game filenames into the playlist or update gamelist.xml
VERSION: 1.4
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
# - "XML" => identify sets and hide Disk 2+ in gamelist.xml
# - "skip" => ignore these platforms entirely (no M3U and no gamelist edits)
$noM3UPlatformMode = "XML"   # <-- set to "skip" to completely ignore those platforms

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
$noM3UNoDisk1Sets              = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

$noM3UNewlyHidden              = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UAlreadyHidden            = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

$noM3UNewlyUnhidden            = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }
$noM3UAlreadyVisible           = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

$noM3UMissingGamelistEntries   = [System.Collections.ArrayList]::new() # PSCustomObject { FullPath; Reason }

# Establish per-platform count buckets for gamelist changes
$gamelistHiddenCounts          = @{}  # platform label -> count newly hidden
$gamelistAlreadyHiddenCounts   = @{}  # platform label -> count already hidden (no change)

$gamelistUnhiddenCounts        = @{}  # platform label -> count newly unhidden
$gamelistAlreadyVisibleCounts  = @{}  # platform label -> count already visible (no change)

# ==================================================================================================
# FUNCTIONS
# ==================================================================================================

function Normalize-M3UText {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) { return @() }

    if ($Text.Length -gt 0 -and [int]$Text[0] -eq 0xFEFF) {
        $Text = $Text.Substring(1)
    }

    $Text = $Text -replace "`r`n", "`n"
    $Text = $Text -replace "`r", "`n"

    return ,($Text -split "`n")
}

function Convert-DiskToSort {
    param([string]$DiskToken)

    if ([string]::IsNullOrWhiteSpace($DiskToken)) { return $null }

    if ($DiskToken -match '^\d+$') { return [int]$DiskToken }

    $romanMap = @{
        'I' = 1;  'II' = 2;  'III' = 3;  'IV' = 4;  'V' = 5
        'VI' = 6; 'VII' = 7; 'VIII' = 8; 'IX' = 9;  'X' = 10
        'XI' = 11; 'XII' = 12; 'XIII' = 13; 'XIV' = 14; 'XV' = 15
        'XVI' = 16; 'XVII' = 17; 'XVIII' = 18; 'XIX' = 19; 'XX' = 20
    }

    $upper = $DiskToken.ToUpperInvariant()
    if ($romanMap.ContainsKey($upper)) { return $romanMap[$upper] }

    if ($upper -match '^[A-Z]$') {
        return ([int][char]$upper[0]) - 64
    }

    return $null
}

function Convert-SideToSort {
    param([string]$SideToken)

    if ([string]::IsNullOrWhiteSpace($SideToken)) { return 0 }

    $c = $SideToken.ToUpperInvariant()[0]
    return ([int][char]$c) - 64
}

function Is-AltTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)[ab]\d*\]$')
}

function Is-DiskNoiseTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)\s*disks?\b')
}

function Is-ExtensionNoiseTag {
    param(
        [Parameter(Mandatory=$true)][string]$Tag,
        [Parameter(Mandatory=$true)][string]$FileName
    )

    if (Is-AltTag $Tag) { return $false }

    $ext = [System.IO.Path]::GetExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($ext)) { return $false }
    $extNoDot = $ext.TrimStart('.')
    if ([string]::IsNullOrWhiteSpace($extNoDot)) { return $false }

    $m = [regex]::Match($Tag, '^\[(?<X>[^\]]+)\]$')
    if (-not $m.Success) { return $false }
    $inner = $m.Groups['X'].Value.Trim()

    return ($inner -ieq $extNoDot)
}

function Clean-BasePrefix {
    param([string]$Prefix)

    if ($null -eq $Prefix) { return "" }

    $p = $Prefix.Trim()
    $p = $p -replace '[\s._-]+$', ''
    $p = $p -replace '\(\s*$', ''

    return $p.Trim()
}

function Write-Phase {
    param([string]$Text)
    Write-Host ""
    Write-Host $Text -ForegroundColor Cyan
}

function Get-AltFallbackChain {
    param([string]$AltKey)

    if ([string]::IsNullOrWhiteSpace($AltKey)) { return @("") }

    $m = [regex]::Match($AltKey, '^\[(?i)(?<L>[ab])(?<N>\d*)\]$')
    if (-not $m.Success) { return @($AltKey, "") }

    $letter = $m.Groups['L'].Value.ToLowerInvariant()
    $num = $m.Groups['N'].Value

    if ([string]::IsNullOrWhiteSpace($num)) {
        return @($AltKey, "")
    }

    return @($AltKey, "[$letter]", "")
}

function Get-NonBangTagsKey {
    param([string]$BaseTagsKey)
    if ([string]::IsNullOrWhiteSpace($BaseTagsKey)) { return "" }
    return ($BaseTagsKey -replace '\[\!\]', '')
}

function Normalize-Alt {
    param([AllowNull()][AllowEmptyString()][string]$Alt)
    if ([string]::IsNullOrWhiteSpace($Alt)) { return $null }
    return $Alt
}

function Get-PlatformRootName {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$ScriptDir
    )

    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $dirFull    = (Resolve-Path -LiteralPath $Directory).Path.TrimEnd('\')

    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    if ($scriptIsRomsRoot) {
        if (-not $dirFull.StartsWith($scriptFull, [System.StringComparison]::OrdinalIgnoreCase)) { return $null }
        $rel = $dirFull.Substring($scriptFull.Length).TrimStart('\')
        if ([string]::IsNullOrWhiteSpace($rel)) { return $null }
        $parts = $rel -split '\\'
        if ($parts.Count -ge 1) { return $parts[0].ToLowerInvariant() }
        return $null
    }

    return $scriptLeaf.ToLowerInvariant()
}

function Get-PlatformRootPath {
    param(
        [Parameter(Mandatory=$true)][string]$ScriptDir,
        [Parameter(Mandatory=$true)][string]$PlatformRootName
    )

    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    if ($scriptIsRomsRoot) {
        return (Join-Path $scriptFull $PlatformRootName)
    }

    return $scriptFull
}

function Get-PlatformCountLabel {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$ScriptDir
    )

    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $dirFull    = (Resolve-Path -LiteralPath $Directory).Path.TrimEnd('\')

    if (-not $dirFull.StartsWith($scriptFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return (Split-Path -Leaf $dirFull).ToUpperInvariant()
    }

    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $rel = $dirFull.Substring($scriptFull.Length).TrimStart('\')
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($rel)) { $parts = $rel -split '\\' }

    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    if ($scriptIsRomsRoot) {
        if ($parts.Count -eq 0) { return $scriptLeaf.ToUpperInvariant() }
        $platform = $parts[0].ToUpperInvariant()
        $subParts = if ($parts.Count -gt 1) { $parts[1..($parts.Count-1)] } else { @() }
        if ($subParts.Count -gt 0) { return ($platform + "\" + ($subParts -join "\")) }
        return $platform
    }

    $platform = $scriptLeaf.ToUpperInvariant()
    if ($parts.Count -gt 0) { return ($platform + "\" + ($parts -join "\")) }
    return $platform
}

function Get-RelativeGamelistPath {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformRootPath,
        [Parameter(Mandatory=$true)][string]$FileFullPath
    )

    try {
        $rootFull = (Resolve-Path -LiteralPath $PlatformRootPath).Path.TrimEnd('\')
        $fileFull = (Resolve-Path -LiteralPath $FileFullPath).Path
    } catch {
        return $null
    }

    if (-not $fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    $rel = $fileFull.Substring($rootFull.Length).TrimStart('\')
    if ([string]::IsNullOrWhiteSpace($rel)) { return $null }

    $rel = $rel -replace '\\', '/'
    return ("./" + $rel)
}

function Get-UniqueGamelistBackupPath {
    param([Parameter(Mandatory=$true)][string]$GamelistPath)

    $dir = Split-Path -Parent $GamelistPath
    $base = Join-Path $dir "gamelist.backup"
    if (-not (Test-Path -LiteralPath $base)) { return $base }

    $i = 1
    while ($true) {
        $p = Join-Path $dir ("gamelist.backup ({0})" -f $i)
        if (-not (Test-Path -LiteralPath $p)) { return $p }
        $i++
    }
}

function Ensure-GamelistLoaded {
    param(
        [Parameter(Mandatory=$true)][string]$PlatformRootLower,
        [Parameter(Mandatory=$true)][string]$PlatformRootPath
    )

    if ($gamelistStateByPlatform.ContainsKey($PlatformRootLower)) {
        return $gamelistStateByPlatform[$PlatformRootLower]
    }

    $gamelistPath = Join-Path $PlatformRootPath "gamelist.xml"

    $state = [PSCustomObject]@{
        RootPath      = $PlatformRootPath
        GamelistPath  = $gamelistPath
        Lines         = $null
        Changed       = $false
        Exists        = (Test-Path -LiteralPath $gamelistPath)
    }

    if ($state.Exists) {
        try {
            $state.Lines = [System.IO.File]::ReadAllLines($gamelistPath)
        } catch {
            $state.Lines = $null
        }
    }

    $gamelistStateByPlatform[$PlatformRootLower] = $state
    return $state
}

function Save-GamelistIfChanged {
    param([Parameter(Mandatory=$true)]$State)

    if (-not $State.Exists -or -not $State.Changed -or $null -eq $State.Lines) { return $false }

    if ($dryRun) { return $false }

    if (-not $gamelistBackupDone.ContainsKey($State.GamelistPath)) {
        $backupPath = Get-UniqueGamelistBackupPath -GamelistPath $State.GamelistPath
        Copy-Item -LiteralPath $State.GamelistPath -Destination $backupPath -Force
        $gamelistBackupDone[$State.GamelistPath] = $true
    }

    $text = ($State.Lines -join [Environment]::NewLine)
    [System.IO.File]::WriteAllText($State.GamelistPath, $text, [System.Text.UTF8Encoding]::new($false))
    return $true
}

function Hide-GameEntriesInGamelist {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)]$Targets,
        [Parameter(Mandatory=$true)][hashtable]$UsedFiles
    )

    $result = [PSCustomObject]@{
        DidWork            = $false
        NewlyHiddenCount   = 0
        AlreadyHiddenCount = 0
        MissingCount       = 0
    }

    if (-not $State.Exists -or $null -eq $State.Lines) {

        foreach ($t in $Targets) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true

            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })

            $result.MissingCount++
        }

        return $result
    }

    $lines = $State.Lines

    foreach ($t in $Targets) {

        $rel = $t.RelPath
        if ([string]::IsNullOrWhiteSpace($rel)) {

            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
            continue
        }

        $found = $false
        $handled = $false

        for ($i = 0; $i -lt $lines.Count; $i++) {

            $line = $lines[$i]
            if ($null -eq $line) { continue }

            $m = [regex]::Match($line, '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
            if (-not $m.Success) { continue }

            $val = $m.Groups['V'].Value
            if ($val -ne $rel) { continue }

            $found = $true
            $indent = $m.Groups['I'].Value

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

            if ($null -ne $hiddenLineIndex) {

                if ($hiddenValue -match '^(?i)true$') {
                    $UsedFiles[$t.FullPath] = $true
                    [void]$noM3UAlreadyHidden.Add([PSCustomObject]@{
                        FullPath = $t.FullPath
                        Reason   = "Already hidden in gamelist.xml"
                    })
                    $result.AlreadyHiddenCount++
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
                    $handled = $true
                }
            }
            else {

                if (-not $dryRun) {

                    $insertLine = ($indent + "<hidden>true</hidden>")

                    $before = @()
                    if ($i -ge 0) { $before = $lines[0..$i] }
                    $after = @()
                    if (($i + 1) -le ($lines.Count - 1)) { $after = $lines[($i + 1)..($lines.Count - 1)] }

                    $lines = @($before + $insertLine + $after)
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
                $i++
            }

            if ($handled) { break }
        }

        if (-not $found) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
        }
    }

    $State.Lines = $lines
    return $result
}

function Unhide-GameEntriesInGamelist {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)]$Targets,
        [Parameter(Mandatory=$true)][hashtable]$UsedFiles
    )

    $result = [PSCustomObject]@{
        DidWork             = $false
        NewlyUnhiddenCount  = 0
        AlreadyVisibleCount = 0
        MissingCount        = 0
    }

    if (-not $State.Exists -or $null -eq $State.Lines) {

        foreach ($t in $Targets) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
        }

        return $result
    }

    $lines = $State.Lines

    foreach ($t in $Targets) {

        $rel = $t.RelPath
        if ([string]::IsNullOrWhiteSpace($rel)) {

            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
            continue
        }

        $found = $false
        $handled = $false

        for ($i = 0; $i -lt $lines.Count; $i++) {

            $line = $lines[$i]
            if ($null -eq $line) { continue }

            $m = [regex]::Match($line, '^(?<I>\s*)<path>\s*(?<V>.*?)\s*</path>\s*$')
            if (-not $m.Success) { continue }

            $val = $m.Groups['V'].Value
            if ($val -ne $rel) { continue }

            $found = $true

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

            if ($null -ne $hiddenLineIndex -and $hiddenValue -match '^(?i)true$') {

                if (-not $dryRun) {
                    $before = @()
                    if ($hiddenLineIndex -gt 0) { $before = $lines[0..($hiddenLineIndex - 1)] }
                    $after = @()
                    if (($hiddenLineIndex + 1) -le ($lines.Count - 1)) { $after = $lines[($hiddenLineIndex + 1)..($lines.Count - 1)] }

                    $lines = @($before + $after)
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
                [void]$noM3UAlreadyVisible.Add([PSCustomObject]@{
                    FullPath = $t.FullPath
                    Reason   = "Already visible in gamelist.xml"
                })
                $result.AlreadyVisibleCount++
                $handled = $true
                break
            }
        }

        if (-not $found) {
            $noM3UMissingGamelistEntryByFullPath[$t.FullPath] = $true
            [void]$noM3UMissingGamelistEntries.Add([PSCustomObject]@{
                FullPath = $t.FullPath
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            $result.MissingCount++
        }
    }

    $State.Lines = $lines
    return $result
}

function Parse-GameFile {
    param(
        [Parameter(Mandatory=$true)][string]$FileName,
        [Parameter(Mandatory=$true)][string]$Directory
    )

    $nameNoExt = $FileName -replace '\.[^\.]+$', ''

    $diskPattern = '(?i)^(?<Prefix>.*?)(?<Sep>[\s._\-\(]|^)(?<Type>disk|disc)(?!s)[\s_]*(?<Disk>\d+|[A-Za-z]|(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX))(?:(?=\s+of\s+\d+)|(?=\s+Side\s+[A-Za-z])|(?=[\s\)\]\._-]|$))(?:\s+of\s+(?<Total>\d+))?(?:\s+Side\s+(?<Side>[A-Za-z]))?(?<After>.*)$'

    $sideOnlyPattern = '(?i)^(?<Prefix>.*?)(?<Sep>[\s._\-\(]|^)Side\s+(?<SideOnly>[A-Za-z])(?=[\s\)\]\._-]|$)(?<After>.*)$'

    $diskMatch = [regex]::Match($nameNoExt, $diskPattern)
    $sideOnlyMatch = $null

    $hasDisk = $diskMatch.Success
    if (-not $hasDisk) {
        $sideOnlyMatch = [regex]::Match($nameNoExt, $sideOnlyPattern)
    }

    if (-not $hasDisk -and (-not $sideOnlyMatch.Success)) {
        return $null
    }

    $prefixRaw  = ""
    $diskToken  = $null
    $totalToken = $null
    $sideToken  = $null
    $after      = ""

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

    $basePrefix = Clean-BasePrefix $prefixRaw
    $afterNorm = $after -replace '^[\)\s]+', ''

    $nameHint = ""
    $beforeBracket = $afterNorm
    $bracketIdx = $beforeBracket.IndexOf('[')
    if ($bracketIdx -ge 0) { $beforeBracket = $beforeBracket.Substring(0, $bracketIdx) }
    $mHint = [regex]::Match($beforeBracket, '^\s*(\([^\)]+\))')
    if ($mHint.Success) { $nameHint = $mHint.Groups[1].Value }

    $bracketTags = @()
    foreach ($m in [regex]::Matches($afterNorm, '\[[^\]]+\]')) {
        $tag = $m.Value
        if (Is-DiskNoiseTag $tag) { continue }
        if (Is-ExtensionNoiseTag -Tag $tag -FileName $FileName) { continue }
        $bracketTags += $tag
    }

    $altTag  = ""
    $baseTags = @()
    foreach ($t in $bracketTags) {
        if (Is-AltTag $t) { $altTag = $t }
        else { $baseTags += $t }
    }

    $baseTagsKey = ($baseTags -join "")

    $diskSort = Convert-DiskToSort $diskToken
    $sideSort = Convert-SideToSort $sideToken

    $totalDisks = $null
    if ($totalToken -match '^\d+$') { $totalDisks = [int]$totalToken }

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

    $wantAlt = Normalize-Alt $AltTag

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

$parsed = @()

Write-Phase "Collecting ROM file data (scanning folders, which might take a while)..."

Get-ChildItem -Path $scriptDir -File -Recurse -Depth 2 | ForEach-Object {

    if ($_.Extension -ieq ".m3u") { return }
    if ($skipFolders -contains $_.Directory.Name.ToLowerInvariant()) { return }

    if ($noM3UPlatformMode -ieq "skip") {

        $plat = Get-PlatformRootName -Directory $_.DirectoryName -ScriptDir $scriptDir
        if ($null -ne $plat) {

            $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
            if ($noM3USetLower -contains $plat.ToLowerInvariant()) { return }
        }
    }

    $p = Parse-GameFile -FileName $_.Name -Directory $_.DirectoryName
    if ($null -ne $p) { $parsed += $p }
}

# ==================================================================================================
# PHASE 2: INDEXING / GROUPING
# ==================================================================================================

Write-Phase "Indexing parsed candidates (grouping titles, tags, variants, etc.)..."

$groupsStrict = $parsed | Group-Object Directory, BasePrefix, BaseTagsKey

$titleIndex = @{}
foreach ($p in $parsed) {
    if (-not $titleIndex.ContainsKey($p.TitleKey)) { $titleIndex[$p.TitleKey] = @() }
    $titleIndex[$p.TitleKey] += $p
}

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

$occupiedPlaylistPaths = @{}
$m3uWrittenPlaylistPaths = @{}
$playlistSignatures = @{}

$suppressedDuplicatePlaylists   = @{} # playlistPath -> collidedWithPath
$suppressedPreExistingPlaylists = @{} # playlistPath -> $true (content identical)
$overwrittenExistingPlaylists   = @{} # playlistPath -> $true (content differed)

$usedFiles = @{}

# ==================================================================================================
# PHASE 3: PROCESS MULTI-DISK GROUPS (M3U PLAYLISTS + NON-M3U GAMELIST HIDING)
# ==================================================================================================

Write-Phase "Processing multi-disk candidates (playlists / gamelist updates)..."

foreach ($group in $groupsStrict) {

    $groupFiles = $group.Group
    $directory  = $groupFiles[0].Directory
    $basePrefix = $groupFiles[0].BasePrefix
    $titleKey   = $groupFiles[0].TitleKey

    $titleFiles = if ($titleIndex.ContainsKey($titleKey)) { $titleIndex[$titleKey] } else { $groupFiles }
    $strictNBKey = $groupFiles[0].BaseTagsKeyNB

    $titleCompatible = @(
        $titleFiles | Where-Object { $_.BaseTagsKeyNB -eq $strictNBKey }
    )

    $altKeys = @(
        ($groupFiles | Select-Object -ExpandProperty AltTag | ForEach-Object { if ($_ -eq $null) { "" } else { $_ } } | Sort-Object -Unique)
    )

    foreach ($altKey in $altKeys) {

        $disk1Roots = @($titleCompatible | Where-Object { $_.DiskSort -eq 1 })

        $totals = @($disk1Roots | Where-Object { $_.TotalDisks -ne $null } | Select-Object -ExpandProperty TotalDisks | Sort-Object -Unique)
        if ($totals.Count -eq 0) { $totals = @($null) }

        foreach ($rootTotal in $totals) {

            $diskTargets = if ($rootTotal) { 1..$rootTotal }
                           else { @($titleCompatible | Select-Object -ExpandProperty DiskSort | Sort-Object -Unique) }

            $playlistFiles = @()

            foreach ($d in $diskTargets) {

                $picked = @()

                $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $altKey -RootTotal $rootTotal

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

                if ($picked.Count -eq 0) {
                    $altChain = Get-AltFallbackChain $altKey
                    foreach ($tryAlt in $altChain) {
                        $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $tryAlt -RootTotal $rootTotal
                        if ($picked.Count -gt 0) { break }
                    }
                }

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

                if ($picked.Count -gt 0) { $playlistFiles += $picked }
            }

            if (@($playlistFiles).Count -lt 2) { continue }

            $uniqueHints = @($playlistFiles | Select-Object -ExpandProperty NameHint | Sort-Object -Unique)
            $useHint = ""
            if ($uniqueHints.Count -eq 1 -and (-not [string]::IsNullOrWhiteSpace($uniqueHints[0]))) { $useHint = $uniqueHints[0] }

            $playlistBase = $basePrefix
            if ($useHint) { $playlistBase += $useHint }

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

            $altsInPlaylist = @($playlistFiles | ForEach-Object { Normalize-Alt $_.AltTag } | Sort-Object -Unique)
            if ($altsInPlaylist.Count -eq 1 -and $altsInPlaylist[0]) {
                $playlistBase += $altsInPlaylist[0]
            }

            $playlistBase = $playlistBase -replace '\s{2,}', ' '
            $playlistBase = $playlistBase -replace '[\s._-]*[\(]*$', ''
            $playlistBase = $playlistBase -replace '\(\s*\)', ''
            $playlistBase = $playlistBase.Trim()
            if ([string]::IsNullOrWhiteSpace($playlistBase)) { continue }

            $platformRoot = Get-PlatformRootName -Directory $directory -ScriptDir $scriptDir

            $isNoM3U = $false
            if ($null -ne $platformRoot) {
                $rootLower = $platformRoot.ToLowerInvariant()
                $noM3USetLower = @($nonM3UPlatforms | ForEach-Object { $_.ToLowerInvariant() })
                $isNoM3U = ($noM3USetLower -contains $rootLower)
            }

            if ($isNoM3U -and $noM3UPlatformMode -ieq "skip") { continue }

            $sorted = $playlistFiles | Sort-Object DiskSort, SideSort

            if ($isNoM3U) {

                $disk1Candidates = @($sorted | Where-Object { $_.DiskSort -eq 1 } | Sort-Object SideSort)
                $hasDisk1 = ($disk1Candidates.Count -gt 0)

                if (-not $hasDisk1) {

                    $platformLower = $platformRoot.ToLowerInvariant()
                    $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
                    $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

                    $targetsU = @()
                    foreach ($sf in $sorted) {
                        $fullFile = Join-Path $sf.Directory $sf.FileName
                        $rel = Get-RelativeGamelistPath -PlatformRootPath $rootPath -FileFullPath $fullFile
                        $targetsU += [PSCustomObject]@{
                            FullPath      = $fullFile
                            RelPath       = $rel
                            PlatformLabel = $platformRoot.ToUpperInvariant()
                        }
                    }

                    $unhideResult = Unhide-GameEntriesInGamelist -State $state -Targets $targetsU -UsedFiles $usedFiles

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

            $setIsIncomplete = $false
            if ($rootTotal) {
                $presentDisks = @($playlistFiles | Select-Object -ExpandProperty DiskSort | Sort-Object -Unique)
                $missingDisksLocal = @()
                foreach ($ed in (1..$rootTotal)) {
                    if (-not ($presentDisks -contains $ed)) { $missingDisksLocal += $ed }
                }
                if ($missingDisksLocal.Count -gt 0) { $setIsIncomplete = $true }
            }

            $thisHasBang = ($groupFiles[0].BaseTagsKey -match '\[\!\]')
            $wantAltNorm = Normalize-Alt $altKey
            if (-not $thisHasBang -and $wantAltNorm) {
                $kBang    = $titleKey + "`0" + $strictNBKey
                $kBangAlt = $titleKey + "`0" + $strictNBKey + "`0" + $wantAltNorm
                if ($bangByTitleNB.ContainsKey($kBang) -and $bangAltByTitleNB.ContainsKey($kBangAlt)) {
                    continue
                }
            }

            $newLines = @($sorted | ForEach-Object { $_.FileName })
            $cleanLines = @(
                $newLines |
                    ForEach-Object { $_.TrimEnd() } |
                    Where-Object { $_ -ne "" }
            )

            $cleanText = ($cleanLines -join "`n")
            $newNorm   = Normalize-M3UText $cleanText

            $playlistPath = if ($isNoM3U) {
                Join-Path $directory "$playlistBase"
            } else {
                Join-Path $directory "$playlistBase.m3u"
            }

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

            $sigParts = @()
            foreach ($sf in $sorted) { $sigParts += (Join-Path $sf.Directory $sf.FileName) }
            $playlistSig = ($sigParts -join "`0")

            if ($playlistSignatures.ContainsKey($playlistSig)) {

                $suppressedDuplicatePlaylists[$playlistPath] = $playlistSignatures[$playlistSig]
                $occupiedPlaylistPaths[$playlistPath] = $true

                foreach ($sf in $sorted) {
                    $full = Join-Path $sf.Directory $sf.FileName
                    $usedFiles[$full] = $true
                }

                continue
            }

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

            $playlistSignatures[$playlistSig] = $playlistPath
            $occupiedPlaylistPaths[$playlistPath] = $true

            if ($isNoM3U) {

                $disk1Candidates = @($sorted | Where-Object { $_.DiskSort -eq 1 } | Sort-Object SideSort)
                if ($disk1Candidates.Count -eq 0) { continue }

                $primaryObj  = $disk1Candidates[0]
                $primaryFull = Join-Path $primaryObj.Directory $primaryObj.FileName
                $usedFiles[$primaryFull] = $true

                if ($setIsIncomplete) {

                    $platformLower = $platformRoot.ToLowerInvariant()
                    $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
                    $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

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

                $toHide = @(
                    $sorted | Where-Object { (Join-Path $_.Directory $_.FileName) -ne $primaryFull }
                )

                if ($toHide.Count -eq 0) {
                    [void]$noM3UPrimaryEntriesOk.Add([PSCustomObject]@{
                        FullPath = $primaryFull
                        Reason   = "Primary entry kept visible"
                    })
                    continue
                }

                $platformLower = $platformRoot.ToLowerInvariant()
                $rootPath = Get-PlatformRootPath -ScriptDir $scriptDir -PlatformRootName $platformLower
                $state = Ensure-GamelistLoaded -PlatformRootLower $platformLower -PlatformRootPath $rootPath

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

                $hideResult = Hide-GameEntriesInGamelist -State $state -Targets $targets -UsedFiles $usedFiles

                $platLabel = $platformRoot.ToUpperInvariant()
                if (-not $gamelistHiddenCounts.ContainsKey($platLabel)) { $gamelistHiddenCounts[$platLabel] = 0 }
                if (-not $gamelistAlreadyHiddenCounts.ContainsKey($platLabel)) { $gamelistAlreadyHiddenCounts[$platLabel] = 0 }
                $gamelistHiddenCounts[$platLabel] += [int]$hideResult.NewlyHiddenCount
                $gamelistAlreadyHiddenCounts[$platLabel] += [int]$hideResult.AlreadyHiddenCount

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

                $platformLabel = $platformRoot.ToUpperInvariant()
                if (-not $platformCounts.ContainsKey($platformLabel)) { $platformCounts[$platformLabel] = 0 }
                $platformCounts[$platformLabel]++
                $totalPlaylistsCreated++

                continue
            }

            if ($setIsIncomplete) { continue }

            if (-not $dryRun) {
                [System.IO.File]::WriteAllText($playlistPath, $cleanText, [System.Text.UTF8Encoding]::new($false))
            }

            $m3uWrittenPlaylistPaths[$playlistPath] = $true

            foreach ($sf in $sorted) {
                $full = Join-Path $sf.Directory $sf.FileName
                $usedFiles[$full] = $true
            }

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

Write-Phase "Finalizing gamelist.xml updates (if any)..."

foreach ($k in $gamelistStateByPlatform.Keys) {
    $st = $gamelistStateByPlatform[$k]
    if ($null -ne $st) {
        Save-GamelistIfChanged -State $st | Out-Null
    }
}

# ==================================================================================================
# PHASE 5: STRUCTURED REPORTING (SPLIT: M3U PLAYLISTS vs GAMELIST(S) UPDATED)
# ==================================================================================================

$anyM3UActivity =
    (@($m3uWrittenPlaylistPaths.Keys).Count -gt 0) -or
    (@($suppressedPreExistingPlaylists.Keys).Count -gt 0) -or
    (@($suppressedDuplicatePlaylists.Keys).Count -gt 0)

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

if (-not $anyM3UActivity -and -not $anyGamelistActivity) {
    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green
    Write-Host "No viable multi-disk files were found to create playlists from." -ForegroundColor Yellow
} else {

    Write-Host ""
    Write-Host "M3U PLAYLISTS" -ForegroundColor Green

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

    if (@($suppressedPreExistingPlaylists.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "SUPPRESSED (PRE-EXISTING PLAYLIST CONTAINED IDENTICAL CONTENT)" -ForegroundColor Green
        $suppressedPreExistingPlaylists.Keys | Sort-Object | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
    }

    if (@($suppressedDuplicatePlaylists.Keys).Count -gt 0) {
        Write-Host ""
        Write-Host "SUPPRESSED (DUPLICATE CONTENT COLLISION DURING THIS RUN)" -ForegroundColor Green
        $suppressedDuplicatePlaylists.GetEnumerator() | Sort-Object Name | ForEach-Object {
            Write-Host "$($_.Key)" -NoNewline -ForegroundColor Gray
            Write-Host " — Identical content collision with $($_.Value)" -ForegroundColor Yellow
        }
    }

    if ($anyGamelistActivity) {
        Write-Host ""

        if ($dryRun) {
            Write-Host "GAMELIST(S) UPDATED (DRY RUN — NO FILES MODIFIED)" -ForegroundColor Green
        } else {
            Write-Host "GAMELIST(S) UPDATED" -ForegroundColor Green
        }

        $changedGamelists = @()
        foreach ($k in $gamelistStateByPlatform.Keys) {
            $st = $gamelistStateByPlatform[$k]
            if ($null -ne $st -and $st.Exists -and $st.Changed -and $null -ne $st.GamelistPath) {
                $changedGamelists += $st.GamelistPath
            }
        }

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

        if (@($noM3UPrimaryEntriesOk).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (PRIMARY ENTRIES KEPT VISIBLE — OK)" -ForegroundColor Green
            $noM3UPrimaryEntriesOk | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        if (@($noM3UPrimaryEntriesIncomplete).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (PRIMARY ENTRIES KEPT VISIBLE — SET INCOMPLETE)" -ForegroundColor Green
            $noM3UPrimaryEntriesIncomplete | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

        if (@($noM3UNoDisk1Sets).Count -gt 0) {
            Write-Host ""
            Write-Host "NON-M3U SETS IDENTIFIED (NO DISK 1 — SET INCOMPLETE)" -ForegroundColor Green
            $noM3UNoDisk1Sets | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

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

        if (@($noM3UAlreadyHidden).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES ALREADY HIDDEN (NO CHANGE)" -ForegroundColor Green
            $noM3UAlreadyHidden | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

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

        if (@($noM3UAlreadyVisible).Count -gt 0) {
            Write-Host ""
            Write-Host "ENTRIES ALREADY VISIBLE (NO CHANGE)" -ForegroundColor Green
            $noM3UAlreadyVisible | Sort-Object FullPath | ForEach-Object {
                Write-Host "$($_.FullPath)" -NoNewline
                Write-Host " — $($_.Reason)" -ForegroundColor Yellow
            }
        }

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

$notInPlaylists = [System.Collections.ArrayList]::new()

$groupsForNotUsed = $parsed | Group-Object TitleKey
foreach ($g in $groupsForNotUsed) {

    $gFiles = $g.Group

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

    $diskSet = @(
        $gFiles |
            Where-Object { $_.DiskSort -ne $null } |
            Select-Object -ExpandProperty DiskSort |
            Sort-Object -Unique
    )
    if ($diskSet.Count -eq 0) { continue }

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

    $altByDisk = @{}
    foreach ($x in $gFiles) {
        if ($null -eq $x.DiskSort) { continue }
        $dsk = [int]$x.DiskSort
        $a = Normalize-Alt $x.AltTag
        if (-not $altByDisk.ContainsKey($dsk)) { $altByDisk[$dsk] = @() }
        if (-not ($altByDisk[$dsk] -contains $a)) { $altByDisk[$dsk] += $a }
    }

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

    $missingDisks = @()
    foreach ($ed in $expectedDisks) {
        if (-not ($diskSet -contains $ed)) { $missingDisks += $ed }
    }

    foreach ($f in $gFiles) {

        $full = Join-Path $f.Directory $f.FileName
        if ($usedFiles.ContainsKey($full)) { continue }

        if ($noM3UMissingGamelistEntryByFullPath.ContainsKey($full)) {
            [void]$notInPlaylists.Add([PSCustomObject]@{
                FullPath = $full
                Reason   = "No entry in gamelist.xml (run gamelist update, scrap the game or manually add it into gamelist.xml)"
            })
            continue
        }

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

        $hasBangInThisFile = ($f.BaseTagsKey -match '\[\!\]')
        $altNorm = Normalize-Alt $f.AltTag
        if (-not $hasBangInThisFile -and $altNorm) {

            $kBang = $f.TitleKey + "`0" + $f.BaseTagsKeyNB
            $kBangAlt = $f.TitleKey + "`0" + $f.BaseTagsKeyNB + "`0" + $altNorm

            if ($bangByTitleNB.ContainsKey($kBang) -and $bangAltByTitleNB.ContainsKey($kBangAlt)) {
                $reason = "Suppressed by [!] rule"
            }
        }

        if ($reason -eq "Unselected during fill" -and $missingDisks.Count -gt 0) {
            $reason = "Missing matching disk"
        }

        if ($reason -eq "Unselected during fill" -and $missingDisks.Count -eq 0 -and $altNorm) {
            $reason = "Alt fallback failed"
        }

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

        if ($reason -eq "Missing matching disk" -and $missingDisks.Count -gt 0) {
            $reason = "Incomplete disk set"
        }

        if ($reason -eq "Incomplete disk set" -and ($missingDisks -contains 1)) {
            $reason = "Incomplete disk set (missing Disk 1)"
        }

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

# Hide this entire count section when TOTAL = 0 (even if keys exist with zero values)
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

# Hide this entire count section when TOTAL = 0 (even if keys exist with zero values)
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

# Hide this entire count section when TOTAL = 0 (even if keys exist with zero values)
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

if (@($notInPlaylists).Count -gt 0) {
    Write-Host ""
    Write-Host "MULTI-DISK FILE SKIP COUNT" -ForegroundColor Green
    Write-Host "TOTAL:" -ForegroundColor White -NoNewline
    Write-Host " $(@($notInPlaylists).Count)"
}

# ==================================================================================================
# PHASE 8: FINAL RUNTIME REPORT
# ==================================================================================================

$elapsed = (Get-Date) - $scriptStart
$totalSeconds = [int][math]::Floor($elapsed.TotalSeconds)

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

Write-Host ""
Write-Host "Runtime:" -ForegroundColor White -NoNewline
Write-Host " $runtimeText"
