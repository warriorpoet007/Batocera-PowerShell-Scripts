<#
PURPOSE: Create .m3u files for each multi-disk game and insert the list of game filenames into the playlist
VERSION: 1.2
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Place this file into the ROMS folder to process all platforms, or in a platform's individual subfolder to process just that one.
- False detections and misses are possible, especially for complex naming structures, but should be rare.
    - I've built intelligence into the script that attempts to determine and provide annotations for a variety of scenarios as to why:
        - The creation of a playlist might have been suppressed
        - A multi-disk file wasn't incorporated into a playlist

BREAKDOWN
- Enumerates ROM/game files starting in the directory the script resides in
     - Scans up to 2 subdirectory levels deep recursively
     - Skips .m3u files during scanning (so it doesn’t treat playlists as input)
     - Skips common media/manual folders (e.g., images, videos, media, manuals, downloaded_*) to reduce false multi-disk detections
- Detects multi-disk candidates by parsing filenames for “designators”
     - A designator is a disk/disc/side marker that indicates a set (case-insensitive), such as:
          - Disk 1, Disc B, Disk II, Disk 2 of 6, Disk 4 Side A
          - Side-only sets like Side A, Side B are supported (treated as Disk 1 with different sides)
     - Supports disk tokens as:
          - Numbers (1, 2, …)
          - Letters (A, B, …)
          - Roman numerals (I … XX) (used for sort normalization)
     - Also recognizes optional patterns like:
          - of N totals (e.g., Disk 2 of 6)
          - Side X paired with a disk marker (e.g., Disk 2 Side B)
- Extracts and interprets bracket tags for grouping and playlist naming
     - Separates tags into:
          - Alt tags like [a], [a2], [b], [b3] (TOSEC-style)
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
- Tracks which disk files were “used”
     - Files included in either written playlists or suppressed playlists are marked “used”
     - Remaining parsed multi-disk candidates that weren’t used are reported as:
          - (POSSIBLE) MULTI-DISK FILES SKIPPED, with a reason such as:
          - incomplete disk set
          - missing matching disk
          - suppressed by [!] preference rule
          - alt fallback issues
          - disk total mismatch issues
- Reporting and summary output
     - Displays (only when non-empty):
          - PLAYLISTS CREATED (only the playlists actually written this run, with overwrite notes)
          - PLAYLISTS SUPPRESSED DUE TO PRE-EXISTING DUPLICATE CONTENT
          - PLAYLISTS SUPPRESSED DUE TO DUPLICATE CONTENT DURING THIS RUN
          - (POSSIBLE) MULTI-DISK FILES SKIPPED
          - PLAYLIST CREATION COUNT(S) per platform and total
          - SUPPRESSED PLAYLIST COUNT(S) with PRE-EXISTING / COLLISIONS / TOTAL
          - MULTI-DISK FILE SKIP COUNT
     - Displays runtime as:
          - "X seconds" (<60s)
          - "M:SS" (<60m)
          - "H:MM:SS" (>=60m)
#>

# --- SCRIPT LOCATION, TIMING, AND SUMMARY COUNTERS ---
# - Resolve script working directory + start timer + init counters
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptStart = Get-Date
$platformCounts = @{}
$totalPlaylistsCreated = 0

# --- CONSOLE OUTPUT SAFETY: BUFFER WIDTH ---
# - Widen buffer to reduce truncation of long paths in output (best-effort)
try {
    $raw = $Host.UI.RawUI
    $size = $raw.BufferSize
    if ($size.Width -lt 300) {
        $raw.BufferSize = New-Object Management.Automation.Host.Size(250, $size.Height)
    }
} catch {
    # Ignore if host doesn't allow resizing (e.g., some terminals)
}

# --- RECURSION FOLDER EXCLUSIONS ---
# - Skip common media/art/manual folders to reduce false “disk” detections
$skipFolders = @(
    'images','videos','media','manuals',
    'downloaded_images','downloaded_videos','downloaded_media','downloaded_manuals'
)

# --- M3U CONTENT NORMALIZATION (FOR STABLE EQUALITY CHECKS) ---
# - Ignores BOM + newline style ONLY (CRLF/LF/CR)
# - PRESERVES trailing spaces and trailing blank lines so they count as differences
function Normalize-M3UText {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) { return @() }

    # Strip UTF-8 BOM if present
    if ($Text.Length -gt 0 -and [int]$Text[0] -eq 0xFEFF) {
        $Text = $Text.Substring(1)
    }

    # Normalize newline style only (CRLF/CR -> LF)
    $Text = $Text -replace "`r`n", "`n"
    $Text = $Text -replace "`r", "`n"

    # Split into lines (PRESERVE empty trailing lines)
    return ,($Text -split "`n")
}

# --- DISK TOKEN SORT NORMALIZATION ---
# - Convert disk token (e.g., 1, A, II) into a sortable integer
function Convert-DiskToSort {
    param([string]$DiskToken)

    if ([string]::IsNullOrWhiteSpace($DiskToken)) { return $null }

    # Numeric disk token
    if ($DiskToken -match '^\d+$') { return [int]$DiskToken }

    # Roman numeral disk token (1..20) — for sorting only
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

# --- SIDE TOKEN SORT NORMALIZATION ---
# - Convert Side token (A/B/...) into a sortable integer (A=1, B=2, ...)
function Convert-SideToSort {
    param([string]$SideToken)
    if ([string]::IsNullOrWhiteSpace($SideToken)) { return 0 }
    $c = $SideToken.ToUpperInvariant()[0]
    return ([int][char]$c) - 64
}

# --- TAG CLASSIFICATION: ALT TAG ---
# - Recognize TOSEC-style alt tags like [a], [a2], [b], [b3]
function Is-AltTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)[ab]\d*\]$')
}

# --- TAG CLASSIFICATION: DISK-NOISE TAG ---
# - Recognize bracketed disk descriptors (e.g., [Disk A]) that should not affect grouping
function Is-DiskNoiseTag {
    param([string]$Tag)
    return ($Tag -match '^\[(?i)\s*disks?\b')
}

# --- PLAYLIST BASE NAME CLEANUP ---
# - Normalize the portion of the filename before the disk designator for stable naming
function Clean-BasePrefix {
    param([string]$Prefix)
    if ($null -eq $Prefix) { return "" }
    $p = $Prefix.Trim()
    $p = $p -replace '[\s._-]+$', ''
    $p = $p -replace '\(\s*$', ''
    return $p.Trim()
}

# --- ALT LOOKUP FALLBACK CHAIN ---
# - Build an ordered list of alt lookups (e.g., [a2] -> [a2], [a], base)
function Get-AltFallbackChain {
    param([string]$AltKey)

    if ([string]::IsNullOrWhiteSpace($AltKey)) { return @("") }

    $m = [regex]::Match($AltKey, '^\[(?i)(?<L>[ab])(?<N>\d*)\]$')
    if (-not $m.Success) { return @($AltKey, "") }

    $letter = $m.Groups['L'].Value.ToLowerInvariant()
    $num = $m.Groups['N'].Value

    if ([string]::IsNullOrWhiteSpace($num)) {
        return @($AltKey, "")
    } else {
        return @($AltKey, "[$letter]", "")
    }
}

# --- TAG RELAXATION: REMOVE [!] ---
# - Derive a stable tag key that ignores the [!] tag only
function Get-NonBangTagsKey {
    param([string]$BaseTagsKey)
    if ([string]::IsNullOrWhiteSpace($BaseTagsKey)) { return "" }
    return ($BaseTagsKey -replace '\[\!\]', '')
}

# --- ALT TOKEN NORMALIZATION ---
# - Convert empty/whitespace alt tokens to $null for consistent comparisons
function Normalize-Alt {
    param([AllowNull()][AllowEmptyString()][string]$Alt)
    if ([string]::IsNullOrWhiteSpace($Alt)) { return $null }
    return $Alt
}

# --- PLATFORM LABELING FOR SUMMARY COUNTS ---
# - Compute a per-platform label based on where the script is run (ROM root vs platform folder)
function Get-PlatformCountLabel {
    param(
        [Parameter(Mandatory=$true)][string]$Directory,
        [Parameter(Mandatory=$true)][string]$ScriptDir
    )

    # Resolve absolute paths for reliable prefix comparisons
    $scriptFull = (Resolve-Path -LiteralPath $ScriptDir).Path.TrimEnd('\')
    $dirFull    = (Resolve-Path -LiteralPath $Directory).Path.TrimEnd('\')

    # Safety: fallback if outside scan root
    if (-not $dirFull.StartsWith($scriptFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return (Split-Path -Leaf $dirFull).ToUpperInvariant()
    }

    # Determine relative path under script root
    $scriptLeaf = (Split-Path -Leaf $scriptFull)
    $rel = $dirFull.Substring($scriptFull.Length).TrimStart('\')
    $parts = @()
    if (-not [string]::IsNullOrWhiteSpace($rel)) { $parts = $rel -split '\\' }

    # Heuristic: script folder named "roms" is treated as ROMS root
    $scriptIsRomsRoot = ($scriptLeaf -match '^(?i)roms$')

    if ($scriptIsRomsRoot) {
        if ($parts.Count -eq 0) { return $scriptLeaf.ToUpperInvariant() } # unlikely
        $platform = $parts[0].ToUpperInvariant()
        $subParts = if ($parts.Count -gt 1) { $parts[1..($parts.Count-1)] } else { @() }
        if ($subParts.Count -gt 0) { return ($platform + "\" + ($subParts -join "\")) }
        return $platform
    } else {
        $platform = $scriptLeaf.ToUpperInvariant()
        if ($parts.Count -gt 0) { return ($platform + "\" + ($parts -join "\")) }
        return $platform
    }
}

# --- FILENAME PARSING / METADATA EXTRACTION ---
# - Parse disk/disc/side patterns + extract tags + build grouping keys
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

    $prefixRaw = ""
    $diskToken = $null
    $totalToken = $null
    $sideToken = $null
    $after = ""

    if ($hasDisk) {
        $prefixRaw = $diskMatch.Groups['Prefix'].Value
        $diskToken = $diskMatch.Groups['Disk'].Value
        if ($diskMatch.Groups['Total'].Success) { $totalToken = $diskMatch.Groups['Total'].Value }
        if ($diskMatch.Groups['Side'].Success)  { $sideToken  = $diskMatch.Groups['Side'].Value }
        $after = $diskMatch.Groups['After'].Value
    } else {
        $prefixRaw = $sideOnlyMatch.Groups['Prefix'].Value
        $diskToken = "1"
        $totalToken = $null
        $sideToken = $sideOnlyMatch.Groups['SideOnly'].Value
        $after = $sideOnlyMatch.Groups['After'].Value
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
        if (-not (Is-DiskNoiseTag $tag)) { $bracketTags += $tag }
    }

    $altTag = ""
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

# --- DISK PICKER: EXACT MATCH WITH OPTIONAL CONSTRAINTS ---
# - Select files for a specific disk number (and optional alt + total constraint)
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

# --- FILE ENUMERATION / PARSING PASS ---
# - Scan up to 2 directory levels deep, skip .m3u, parse files that look like multi-disk candidates
$parsed = @()
Get-ChildItem -Path $scriptDir -File -Recurse -Depth 2 | ForEach-Object {
    if ($_.Extension -ieq ".m3u") { return }
    if ($skipFolders -contains $_.Directory.Name.ToLowerInvariant()) { return }
    $p = Parse-GameFile -FileName $_.Name -Directory $_.DirectoryName
    if ($null -ne $p) { $parsed += $p }
}

# --- STRICT GROUPING (DIRECTORY + TITLE + BASE TAGS) ---
$groupsStrict = $parsed | Group-Object Directory, BasePrefix, BaseTagsKey

# --- TITLE INDEX (DIRECTORY + TITLE) ---
$titleIndex = @{}
foreach ($p in $parsed) {
    if (-not $titleIndex.ContainsKey($p.TitleKey)) { $titleIndex[$p.TitleKey] = @() }
    $titleIndex[$p.TitleKey] += $p
}

# --- PRE-SCAN: [!] PRESENCE BY TITLE + ALT ---
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

# --- PLAYLIST PATH OCCUPANCY TRACKING (THIS RUN) ---
# - Tracks any playlist path claimed during this run (written OR suppressed) to prevent [alt] collisions
$occupiedPlaylistPaths = @{}

# --- PLAYLISTS ACTUALLY WRITTEN (THIS RUN) ---
# - Tracks only playlists actually written (new or overwritten) for reporting under PLAYLISTS CREATED
$writtenPlaylistPaths = @{}

# --- PLAYLIST DUPLICATE-CONTENT TRACKING (DURING THIS RUN) ---
$playlistSignatures = @{}  # key: signature string -> first playlist path

# --- REPORTING: DUPLICATE-CONTENT SUPPRESSED (DURING THIS RUN) ---
$suppressedDuplicatePlaylists = @{}  # key: suppressed playlist path -> original playlist path

# --- REPORTING: PRE-EXISTING DUPLICATE PLAYLISTS ---
$suppressedPreExistingPlaylists = @{}  # key: playlist path -> $true

# --- REPORTING: OVERWRITTEN EXISTING PLAYLISTS ---
$overwrittenExistingPlaylists = @{}    # key: playlist path -> $true

# --- USED FILE TRACKING ---
# - Mark full paths that were actually included in a written OR suppressed playlist
$usedFiles = @{}

# --- MAIN: PLAYLIST GENERATION ---
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

                # 1) Strict: exact alt match within strict group
                $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $altKey -RootTotal $rootTotal

                # 1.5) Conservative: if building base (no-alt) playlist, allow a single unambiguous alt disk
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

                # 2) Relaxed: exact alt match in compatible title set (only [!] can differ)
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

                # 3) Strict: alt fallback chain within strict group
                if ($picked.Count -eq 0) {
                    $altChain = Get-AltFallbackChain $altKey
                    foreach ($tryAlt in $altChain) {
                        $picked = Select-DiskEntries -Files $groupFiles -DiskNumber $d -AltTag $tryAlt -RootTotal $rootTotal
                        if ($picked.Count -gt 0) { break }
                    }
                }

                # 4) Relaxed: alt fallback chain in compatible title set (only [!] can differ)
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

            # Conservative suppression: [!] variant preferred
            $thisHasBang = ($groupFiles[0].BaseTagsKey -match '\[\!\]')
            $wantAltNorm = Normalize-Alt $altKey
            if (-not $thisHasBang -and $wantAltNorm) {
                $kBang = $titleKey + "`0" + $strictNBKey
                $kBangAlt = $titleKey + "`0" + $strictNBKey + "`0" + $wantAltNorm
                if ($bangByTitleNB.ContainsKey($kBang) -and $bangAltByTitleNB.ContainsKey($kBangAlt)) {
                    continue
                }
            }

            # Optional hint naming
            $uniqueHints = @($playlistFiles | Select-Object -ExpandProperty NameHint | Sort-Object -Unique)
            $useHint = ""
            if ($uniqueHints.Count -eq 1 -and (-not [string]::IsNullOrWhiteSpace($uniqueHints[0]))) { $useHint = $uniqueHints[0] }

            # Playlist base name
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

            # Base target path
            $playlistPath = Join-Path $directory "$playlistBase.m3u"

            # --- PATH COLLISION HANDLING (SAME RUN) ---
            # - If a path is already occupied (written OR suppressed) in this run, append [alt], [alt2], ...
            if ($occupiedPlaylistPaths.ContainsKey($playlistPath)) {
                $altIndex = 1
                do {
                    $suffix = if ($altIndex -eq 1) { "[alt]" } else { "[alt$altIndex]" }
                    $playlistPath = Join-Path $directory "$playlistBase$suffix.m3u"
                    $altIndex++
                } while ($occupiedPlaylistPaths.ContainsKey($playlistPath))
            }

            # Sort final playlist entries
            $sorted = $playlistFiles | Sort-Object DiskSort, SideSort

            # Build playlist content (filenames only)
            $newLines = @($sorted | ForEach-Object { $_.FileName })

            # Clean output lines (what we'll actually write)
            $cleanLines = @(
                $newLines |
                    ForEach-Object { $_.TrimEnd() } |
                    Where-Object { $_ -ne "" }
            )

            # (1) Compute the normalized representation from EXACTLY what we would write
            $cleanText = ($cleanLines -join "`n")
            $newNorm   = Normalize-M3UText $cleanText

            # --- DUPLICATE-CONTENT SUPPRESSION DURING THIS RUN ---
            # - Use a stable signature (full paths of disk files) so identical lists don't re-emit
            $sigParts = @()
            foreach ($sf in $sorted) { $sigParts += (Join-Path $sf.Directory $sf.FileName) }
            $playlistSig = ($sigParts -join "`0")

            if ($playlistSignatures.ContainsKey($playlistSig)) {
                $suppressedDuplicatePlaylists[$playlistPath] = $playlistSignatures[$playlistSig]
                $occupiedPlaylistPaths[$playlistPath] = $true

                # Mark members as used so skipped report doesn't inflate
                foreach ($sf in $sorted) {
                    $full = Join-Path $sf.Directory $sf.FileName
                    $usedFiles[$full] = $true
                }

                continue
            }

            # --- PRE-EXISTING DUPLICATE-CONTENT SUPPRESSION (ON-DISK FILE COMPARE) ---
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
                } else {
                    $overwrittenExistingPlaylists[$playlistPath] = $true
                }
            }

            # Record signature now that we know it's unique (or being overwritten intentionally)
            $playlistSignatures[$playlistSig] = $playlistPath
            $occupiedPlaylistPaths[$playlistPath] = $true

            # --- WRITE PLAYLIST FILE (CLEAN OUTPUT) ---
            # (2) Ensure output never has trailing spaces per line and never has blank lines at EOF (or anywhere)
            # (3) Write exactly the clean text with UTF-8 (no BOM) and no trailing newline
            [System.IO.File]::WriteAllText($playlistPath, $cleanText, [System.Text.UTF8Encoding]::new($false))

            $writtenPlaylistPaths[$playlistPath] = $true

            # Mark used files
            foreach ($sf in $sorted) {
                $full = Join-Path $sf.Directory $sf.FileName
                $usedFiles[$full] = $true
            }

            # Update counters (overwrites still count as "created/written this run")
            $platformLabel = Get-PlatformCountLabel -Directory $directory -ScriptDir $scriptDir
            if (-not $platformCounts.ContainsKey($platformLabel)) { $platformCounts[$platformLabel] = 0 }
            $platformCounts[$platformLabel]++
            $totalPlaylistsCreated++
        }
    }
}

# --- REPORT: PLAYLISTS CREATED ---
# - Print only playlists actually written this run, with overwrite notes if applicable
if (@($writtenPlaylistPaths.Keys).Count -gt 0) {
    Write-Host ""
    Write-Host "PLAYLISTS CREATED" -ForegroundColor Green
    $writtenPlaylistPaths.Keys | Sort-Object | ForEach-Object {
        $p = $_
        if ($overwrittenExistingPlaylists.ContainsKey($p)) {
            Write-Host "$p" -NoNewline
            Write-Host " — Overwrote existing playlist that contained content discrepancy" -ForegroundColor Yellow
        } else {
            Write-Host $p
        }
    }
}

# --- REPORT: PRE-EXISTING IDENTICAL-CONTENT PLAYLISTS SUPPRESSED ---
if (@($suppressedPreExistingPlaylists.Keys).Count -gt 0) {
    Write-Host ""
    Write-Host "PLAYLISTS SUPPRESSED (PRE-EXISTING PLAYLIST CONTAINED IDENTICAL CONTENT)" -ForegroundColor Green
    $suppressedPreExistingPlaylists.Keys | Sort-Object | ForEach-Object { Write-Host $_ -ForegroundColor Gray}
}

# --- REPORT: DUPLICATE-CONTENT PLAYLISTS SUPPRESSED (DURING THIS RUN) ---
if (@($suppressedDuplicatePlaylists.Keys).Count -gt 0) {
    Write-Host ""
    Write-Host "PLAYLISTS SUPPRESSED (ANOTHER PLAYLIST CREATED DURING THIS RUN CONTAINED IDENTICAL CONTENT)" -ForegroundColor Green
    $suppressedDuplicatePlaylists.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host "$($_.Key)" -NoNewline -ForegroundColor Gray
        Write-Host " — Identical content collision with $($_.Value)" -ForegroundColor Yellow
    }
}

# --- REPORT: POSSIBLE MULTI-DISK FILES NOT WRITTEN ---
$notInPlaylists = @()

$groupsForNotUsed = $parsed | Group-Object TitleKey, BaseTagsKeyNB
foreach ($g in $groupsForNotUsed) {

    $gFiles = $g.Group
    if (@($gFiles).Count -lt 2) { continue }

    $diskSet = @($gFiles | Where-Object { $_.DiskSort -ne $null } | Select-Object -ExpandProperty DiskSort | Sort-Object -Unique)
    if ($diskSet.Count -eq 0) { continue }

    $maxTotal = ($gFiles |
        Where-Object { $_.TotalDisks -ne $null } |
        Select-Object -ExpandProperty TotalDisks |
        Sort-Object -Descending |
        Select-Object -First 1)

    $expectedDisks = @()
    if ($null -ne $maxTotal -and $maxTotal -ne "") { $expectedDisks = 1..([int]$maxTotal) }
    else { $expectedDisks = $diskSet }

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
        if (-not $usedFiles.ContainsKey($full)) {

            $reason = "Unselected during fill"

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

            if ($reason -eq "Missing matching disk" -and $maxTotal) {
                $reason = "Incomplete disk set"
            }

            if ($reason -eq "Unselected during fill" -and
                $null -ne $minTotal -and $null -ne $maxTotalLocal -and
                $minTotal -ne $maxTotalLocal -and
                $f.DiskSort -gt [int]$minTotal) {

                $reason += " due to disk total mismatch"
            }

            $notInPlaylists += [PSCustomObject]@{
                FullPath = $full
                Reason   = $reason
            }
        }
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

# --- FINAL REPORTING: RUNTIME ---
# - Display runtime as: "X seconds" (<60s), "M:SS" (<60m), or "H:MM:SS" (>=60m)
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
