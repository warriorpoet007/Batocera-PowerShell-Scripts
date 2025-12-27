<#
AUTHOR
Devin Kelley, Distant Thunderworks LLC

PURPOSE
Create .m3u files for each multi-disk game and insert the list of game filenames into the playlist

FUNCTIONAL BREAKDOWN
- Enumerates a list of the files starting in the directory the script resides in
    Scans up to 2 subdirectory levels deep recursively
- Filters the file list by those with "disk" or "disc" in the filename
- Filters the list by those with a number or letter designator after "disk" or "disc" at end of filename
    i.e., ignore all other instances of "disk" or "disc" character strings in filenames
- Creates an .m3u file from the base name (i.e., before the disk/disc designator and filename extension)
    Places the .m3u file into the same folder as the game files themselves
    Replaces any existing .m3u files with the same name
- Inserts the list of filenames meeting the above criteria into the newly created .m3u file

CRITERIA
- Filenames for a given game must be identical, with the only differentiator being the disk/disc designator
- The designator must be the last character in the filename before the period and file extension
- Filenames for a given game must share the same file extension (e.g., .zip, .iso, .chd, .bin, etc.)
#>

# Retrieve the directory where the script resides
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Strict Disc/Disk regex (requires number or letter designator to be at end of file)
# - Base = filename before Disk/Disc
# - Type = Disk/Disc
# - Disk = number (1 or more digits) or single letter
# - Must be immediately before the dot (extension)
$diskRegex = '^(?<Base>.*?)[\s._-]*(?<Type>disc|disk)[\s._-]*(?<Disk>\d+|[A-Za-z])\.[^\.]+$'

$parsedFiles = @()

# Enumerate files up to 2 levels deep
Get-ChildItem -Path $scriptDir -File -Recurse -Depth 2 | ForEach-Object {

    $match = [regex]::Match($_.Name, $diskRegex, 'IgnoreCase')

    # Skip non-matching files
    if (-not $match.Success) { return }

    $diskValue = $match.Groups['Disk'].Value

    # Sorting value: numbers as integers, letters as A=1, B=2...
    if ($diskValue -match '^\d+$') {
        $sortValue = [int]$diskValue
    }
    else {
        $sortValue = [int][char]($diskValue.ToUpper())[0] - 64
    }

    $parsedFiles += [PSCustomObject]@{
        FileName  = $_.Name
        Directory = $_.DirectoryName
        BaseName  = ($match.Groups['Base'].Value).Trim()
        DiskSort  = $sortValue
    }
}

# Group by directory + base name
$groups = $parsedFiles | Group-Object Directory, BaseName

foreach ($group in $groups) {

    $groupFiles = $group.Group
    $directory  = $groupFiles[0].Directory

    # Establish playlist name without file extension
    $playlistBase = [System.IO.Path]::GetFileNameWithoutExtension(
        ($groupFiles[0].BaseName -replace '[\s._-]+$', '').Trim()
    )

    $playlistPath = Join-Path $directory "$playlistBase.m3u"

    # Sort discs by numeric/letter order
    $sortedFiles = $groupFiles | Sort-Object DiskSort

    # Write game filename list in M3U file
    $fileList = $sortedFiles | ForEach-Object { $_.FileName }

    Set-Content -Path $playlistPath -Value $fileList -Encoding UTF8

    Write-Host "Created playlist: $playlistPath"
}