<#
AUTHOR: Devin Kelley, Distant Thunderworks LLC
VERSION: 1.1

PURPOSE
Create .m3u files for each multi-disk game and insert the list of game filenames into the playlist

FUNCTIONAL BREAKDOWN
- Enumerates a list of the files starting in the directory the script resides in
    Scans up to 2 subdirectory levels deep recursively
- Filters the file list by those with "disk" or "disc" in the filename
- Filters the list by those with a number or letter designator after "disk" or "disc" at the end of the filename
    Accommodates a space or underscore between "disk/disc" and the number or letter designator
    Accommodates a closing parenthesis ")" between the number/letter designator and the filename extension
    Accommodates filenames that include a numbered disk of total disks (e.g., Disk 1 of 4)
    Accommodates filenames that include a Side A and B
    Skips common media subfolders (e.g., media, images, video, manuals, downloaded_images, etc.)
- Creates an .m3u file from the base name (i.e., before the disk/disc designator and filename extension)
    Places the .m3u file into the same folder as the game files themselves
    Replaces any existing .m3u files with the same name
- Inserts the list of filenames meeting the designated criteria into the newly created .m3u file

NOTES
- Filenames for a given game must be identical, with the only differentiator being the designator
    If not, separate M3U files will be created for each differentiating disk name
- The designator must be the last character string in the filename before the period and file extension
    With an exemption for a ")" between them
-False detections may be possible, although should be very rare
    You may need to delete .m3u files you know you don't need if the output indicates any
#>

# Retrieve the directory where the script resides
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Folder names to skip entirely (case-insensitive, exact match)
$skipFolders = @(
    'images',
    'videos',
    'media',
    'manuals',
    'downloaded_images',
    'downloaded_videos',
    'downloaded_media',
    'downloaded_manuals'
)

# Strict Disc/Disk regex including the following criteria:
# - Base = filename before Disk/Disc
# - Disk = number (1 or more digits) or single letter
# - Must be immediately before the dot (extension)
# - Allows only spaces or underscores between Disk/Disc and number/letter
# - Optional "Side X" after the number/letter
# - Optional closing parenthesis ")" after number/letter or side
# - Supports "N of M" style designators
# - Supports no-space variants like "Disk1", "DiscA"
$diskRegex = '^(?<Base>.*?)[\s._-]*(?<Type>disc|disk)[\s_]*(?<Disk>\d+(?=\s+of\s+\d+)|\d+|[A-Za-z])(\s+of\s+\d+)?(\s+Side\s+(?<Side>[A-Za-z]))?\)?\.[^\.]+$'

$parsedFiles = @()

# Enumerate files up to 2 levels deep
Get-ChildItem -Path $scriptDir -File -Recurse -Depth 2 | ForEach-Object {

    # Skip files in excluded folders (exact folder name match, case-insensitive)
    if ($skipFolders -contains $_.Directory.Name.ToLowerInvariant()) {
        return
    }

    $match = [regex]::Match($_.Name, $diskRegex, 'IgnoreCase')
    if (-not $match.Success) { return }

    $diskValue = $match.Groups['Disk'].Value
    $sideValue = $match.Groups['Side'].Value

    # Convert disk value to sortable number
    if ($diskValue -match '^\d+$') {
        $diskSort = [int]$diskValue
    } else {
        $diskSort = [int][char]($diskValue.ToUpper())[0] - 64
    }

    # Convert side to sortable number
    if ($sideValue) {
        $sideSort = [int][char]($sideValue.ToUpper())[0] - 64
    } else {
        $sideSort = 0
    }

    $parsedFiles += [PSCustomObject]@{
        FileName  = $_.Name
        Directory = $_.DirectoryName
        BaseName  = ($match.Groups['Base'].Value).Trim()
        DiskSort  = $diskSort
        SideSort  = $sideSort
    }
}

# Group by directory + base name
$groups = $parsedFiles | Group-Object Directory, BaseName

foreach ($group in $groups) {

    $groupFiles = $group.Group
    $directory  = $groupFiles[0].Directory

    # Establish playlist name and modify it by:
        # Removing any trailing spaces, dots, underscores, or hyphens
        # Removing any trailing opening parentheses
        # Removing any remaining leading or trailing whitespace
    $playlistBase = ($groupFiles[0].BaseName -replace '[\s._-]*[\(]*$', '').Trim()

    # Skip output if playlist name is null or empty
    if ([string]::IsNullOrWhiteSpace($playlistBase)) { continue }

    $playlistPath = Join-Path $directory "$playlistBase.m3u"

    # Sort by Disk then Side
    $sortedFiles = $groupFiles | Sort-Object DiskSort, SideSort

    # Write playlist
    $fileList = $sortedFiles | ForEach-Object { $_.FileName }
    $fileList | Out-File -LiteralPath $playlistPath -Encoding UTF8 -Force

    Write-Host "Created playlist: $playlistPath"
}

