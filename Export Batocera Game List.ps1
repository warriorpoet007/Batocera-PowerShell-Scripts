<#
PURPOSE: Export a list of games/apps for Batocera platform folders by reading each platform's gamelist.xml
VERSION: 1.3
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Place this file into the ROMS folder to process all platforms, or in a platform's individual subfolder to process just that one.

- This script does NOT modify any files; it only reads gamelist.xml and produces a CSV report.

- Output CSV is written to the same directory where the PS1 script resides.
    - If "Game List.csv" already exists in the output folder, it is deleted and replaced with a newly generated file.

IMPORTANT NOTE:
- For non-M3U multi-disk games, they must each have the same name in gamelist.xml, as this is what's used by the script to group them
    - Generate Batocera Playlists.ps1 does attempt to do this by populating <name> in gamelist.xml, but this note is here for awareness
    - In the example below, a game with three disks/filenames/paths are all part of the same game with the same name:

        <path>./Game ROM Image (Disk 1).chd</path>
        <name>Name Of the Game</name>

        <path>./Game ROM Image (Disk 2).chd</path>
        <name>Name Of the Game</name>

        <path>./Game ROM Image (Disk 3).chd</path>
        <name>Name Of the Game</name>

BREAKDOWN:
- Multi-disk detection is inferred as:
    - Multi-M3U: the visible entry’s <path> ends in .m3u
    - Multi-XML: the visible entry has 1+ additional entries with the same group key where <hidden>true</hidden> is set
        - This is tailored to be run after first running the Generate Batocera Playlists.ps1 script
        - Ensure that each disk listed in gamelist.xml is tagged with the same <name>Name Of the Game</name>
    - Single: neither of the above

- Gamelist.xml Repair:
    - Some gamelist.xml files in the wild can be malformed (mismatched tags, partial writes, etc.).
    - This script will attempt a normal XML parse first, and if that fails it will fall back to a "salvage mode"
      that extracts <game>...</game> blocks and parses them individually.

- Identifies non-game ROMs via entries that include "ZZZ(notgame):" in the <name> tag in gamelist.xml
    - Removes "ZZZ(notgame):" before it writes the Title
    - Designates this ROM type with a "Game?" column in the CSV file, with a "No" (listings with "Yes" are games)

- XMLState column:
    - Normal: gamelist.xml parsed cleanly as a complete XML document
    - Malformed: gamelist.xml was malformed; entries were extracted by parsing <game> fragments

- Progress / phase output:
    - Prints only major phase steps
    - If running from ROMS root (multi-platform mode), prints per-platform start + finished lines
    - Always prints a final "finished" summary

- Determines runtime mode based on where the script is located:
    - If a gamelist.xml exists in the script directory, treat it as a single-platform run
    - Otherwise, treat the script directory as ROMS root and scan all first-level subfolders for gamelist.xml

- Reads each discovered gamelist.xml and extracts per-game fields:
    - Name (from <name>, with filename fallback if <name> is missing)
    - Path (from <path>)
    - Hidden (from <hidden>, if present; defaults to false)

- Uses a stable group key for multi-disk inference:
    - Primary: <name>
    - Fallback: <path> when <name> is missing/blank (prevents unrelated entries from collapsing into one group)

- Groups entries by the group key to infer multi-disk sets and produce a single row per visible entry

- Generates additional derived columns:
    - Title (Title Case derived from the name; uses safe fallbacks when name is missing)
    - PlatformName (friendly platform name derived from the platform folder)
    - Manufacturer (supplemental platform note derived from the platform mapping)
    - EntryType (Single, Multi-M3U, Multi-XML)
    - DiskCount:
        - Multi-M3U: counts non-empty, non-comment lines in the .m3u file
        - Multi-XML: total entries in the group (visible + hidden)
        - Single: 1
    - XMLState (Normal or Malformed)

- Deletes any existing "Game List.csv" in the output folder and writes a newly generated "Game List.csv"
#>

# ==================================================================================================
# SCRIPT STARTUP: PATHS AND TARGET DISCOVERY
# ==================================================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------------------------------
# Script location and run roots
# PURPOSE:
# - Ensure behavior is based on where the script resides:
#     - ROMS root mode when script is placed in ROMS
#     - Single-platform mode when script is placed in a platform folder
# - Ensure output CSV is written next to the script
# --------------------------------------------------------------------------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$startDir  = $scriptDir

# --------------------------------------------------------------------------------------------------
# Runtime tracking
# PURPOSE:
# - Provide a consistent runtime report using the same formatting rules as the Generate List script:
#     - <60s  => "X seconds"
#     - <60m  => "M:SS"
#     - >=60m => "H:MM:SS"
# --------------------------------------------------------------------------------------------------
$__runtimeStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

function Format-ElapsedRuntime {
  param([Parameter(Mandatory=$true)][TimeSpan]$Elapsed)

  if ($Elapsed.TotalSeconds -lt 60) {
    $sec = [int][math]::Floor($Elapsed.TotalSeconds)
    return ("{0} seconds" -f $sec)
  }

  if ($Elapsed.TotalMinutes -lt 60) {
    $min = [int][math]::Floor($Elapsed.TotalMinutes)
    $sec = [int]$Elapsed.Seconds
    return ("{0}:{1:00}" -f $min, $sec)
  }

  $hrs = [int][math]::Floor($Elapsed.TotalHours)
  $min = [int]$Elapsed.Minutes
  $sec = [int]$Elapsed.Seconds
  return ("{0}:{1:00}:{2:00}" -f $hrs, $min, $sec)
}

function Write-RuntimeReport {
  param([switch]$Stop)

  if ($null -eq $__runtimeStopwatch) { return }

  try {
    if ($Stop -and $__runtimeStopwatch.IsRunning) { $__runtimeStopwatch.Stop() }
    $rt = Format-ElapsedRuntime -Elapsed $__runtimeStopwatch.Elapsed
    Write-Host ("Runtime: {0}" -f $rt) -ForegroundColor DarkYellow
  } catch {
    # Intentionally ignore runtime reporting failures
  }
}

# --------------------------------------------------------------------------------------------------
# Phase output helper
# PURPOSE:
# - Emit clear, consistent phase/progress messages to the console
# NOTES:
# - Intentionally not chatty: major steps only + optional per-platform lines in ROMS root mode
# --------------------------------------------------------------------------------------------------
function Write-Phase {
  param([string]$Message)
  Write-Host ""
  Write-Host $Message -ForegroundColor Cyan
}

# --------------------------------------------------------------------------------------------------
# Runtime mode determination
# PURPOSE:
# - Decide whether we are scanning a single platform or the ROMS root
# NOTES:
# - Single-platform mode: script directory contains gamelist.xml
# - ROMS root mode: script directory does not contain gamelist.xml
# --------------------------------------------------------------------------------------------------
$localGamelistPath     = Join-Path $startDir 'gamelist.xml'
$isSinglePlatformMode  = (Test-Path -LiteralPath $localGamelistPath)
$isRomsRootMode        = (-not $isSinglePlatformMode)

Write-Phase "Starting export..."

# ==================================================================================================
# USER CONFIGURATION: PLATFORM NAME TRANSLATIONS
# ==================================================================================================

# Folder -> Platform + Manufacturer mapping (extend as needed)
# - Platform: friendly platform name (no parentheses)
# - Manufacturer: supplemental note (manufacturer, port, etc.) without parentheses; may be ""
$PlatformMap = @{
  '3do'            = @{ Platform = '3DO'; Manufacturer = 'Panasonic' }
  '3ds'            = @{ Platform = 'Nintendo 3DS'; Manufacturer = 'Nintendo' }
  'abuse'          = @{ Platform = 'Abuse SDL'; Manufacturer = 'Port' }
  'adam'           = @{ Platform = 'Coleco Adam'; Manufacturer = 'Coleco' }
  'advision'       = @{ Platform = 'Adventure Vision'; Manufacturer = 'Entex' }
  'amiga1200'      = @{ Platform = 'Amiga 1200/AGA'; Manufacturer = 'Commodore' }
  'amiga500'       = @{ Platform = 'Amiga 500/OCS/ECS'; Manufacturer = 'Commodore' }
  'amigacd32'      = @{ Platform = 'Amiga CD32'; Manufacturer = 'Commodore' }
  'amigacdtv'      = @{ Platform = 'Commodore CDTV'; Manufacturer = 'Commodore' }
  'amstradcpc'     = @{ Platform = 'Amstrad CPC'; Manufacturer = 'Amstrad' }
  'apfm1000'       = @{ Platform = 'APF-MP1000/MP-1000/M-1000'; Manufacturer = 'APF Electronics Inc.' }
  'apple2'         = @{ Platform = 'Apple II'; Manufacturer = 'Apple' }
  'apple2gs'       = @{ Platform = 'Apple IIGS'; Manufacturer = 'Apple' }
  'arcadia'        = @{ Platform = 'Arcadia 2001'; Manufacturer = 'Emerson Radio' }
  'archimedes'     = @{ Platform = 'Archimedes'; Manufacturer = 'Acorn Computers' }
  'arduboy'        = @{ Platform = 'Arduboy'; Manufacturer = 'Arduboy' }
  'astrocde'       = @{ Platform = 'Astrocade'; Manufacturer = 'Bally/Midway' }
  'atari2600'      = @{ Platform = 'Atari 2600/VCS'; Manufacturer = 'Atari' }
  'atari5200'      = @{ Platform = 'Atari 5200'; Manufacturer = 'Atari' }
  'atari7800'      = @{ Platform = 'Atari 7800'; Manufacturer = 'Atari' }
  'atari800'       = @{ Platform = 'Atari 800'; Manufacturer = 'Atari' }
  'atarist'        = @{ Platform = 'Atari ST'; Manufacturer = 'Atari' }
  'atom'           = @{ Platform = 'Atom'; Manufacturer = 'Acorn Computers' }
  'atomiswave'     = @{ Platform = 'Sammy Atomiswave'; Manufacturer = 'Sammy' }
  'bbc'            = @{ Platform = 'BBC Micro/Master/Archimedes'; Manufacturer = 'Acorn Computers' }
  'bennugd'        = @{ Platform = 'BennuGD'; Manufacturer = 'Game Development Suite' }
  'boom3'          = @{ Platform = 'Doom 3'; Manufacturer = 'Port' }
  'camplynx'       = @{ Platform = 'Camputers Lynx'; Manufacturer = 'Camputers' }
  'cannonball'     = @{ Platform = 'Cannonball'; Manufacturer = 'Port' }
  'casloopy'       = @{ Platform = 'Casio Loopy'; Manufacturer = 'Casio' }
  'catacombgl'     = @{ Platform = 'Catacomb GL'; Manufacturer = 'Port' }
  'cavestory'      = @{ Platform = 'Cave Story'; Manufacturer = 'Port' }
  'c128'           = @{ Platform = 'Commodore 128'; Manufacturer = 'Commodore' }
  'c20'            = @{ Platform = 'Commodore VIC-20/VC-20'; Manufacturer = 'Commodore' }
  'c64'            = @{ Platform = 'Commodore 64'; Manufacturer = 'Commodore' }
  'cdi'            = @{ Platform = 'Compact Disc Interactive/CD-i'; Manufacturer = 'Philips, et al.' }
  'cdogs'          = @{ Platform = 'C-Dogs'; Manufacturer = 'Port' }
  'cgenius'        = @{ Platform = 'Commander Genius (Commander Keen and Cosmos the Cosmic Adventure)'; Manufacturer = 'Port' }
  'channelf'       = @{ Platform = 'Fairchild Channel F'; Manufacturer = 'Fairchild' }
  'chihiro'        = @{ Platform = 'Chihiro'; Manufacturer = 'Sega' }
  'coco'           = @{ Platform = 'TRS-80/Color Computer'; Manufacturer = 'Tandy/RadioShack' }
  'colecovision'   = @{ Platform = 'ColecoVision'; Manufacturer = 'Coleco' }
  'commanderx16'   = @{ Platform = 'Commander X16'; Manufacturer = 'David Murray' }
  'corsixth'       = @{ Platform = 'CorsixTH (Theme Hospital)'; Manufacturer = 'Port' }
  'cplus4'         = @{ Platform = 'Commodore Plus/4'; Manufacturer = 'Commodore' }
  'crvision'       = @{ Platform = 'CreatiVision/Educat 2002/Dick Smith Wizzard/FunVision'; Manufacturer = 'VTech' }
  'daphne'         = @{ Platform = 'DAPHNE Laserdisc'; Manufacturer = 'Various' }
  'devilutionx'    = @{ Platform = 'DevilutionX (Diablo/Hellfire)'; Manufacturer = 'Port' }
  'dice'           = @{ Platform = 'Discrete Integrated Circuit Emulator'; Manufacturer = 'Various' }
  'dolphin'        = @{ Platform = 'Dolphin'; Manufacturer = 'GameCube/Wii Emulator' }
  'dos'            = @{ Platform = 'DOSbox'; Manufacturer = 'Peter Veenstra/Sjoerd van der Berg' }
  'dreamcast'      = @{ Platform = 'Dreamcast'; Manufacturer = 'Sega' }
  'dxx-rebirth'    = @{ Platform = 'DXX Rebirth (Descent/Descent 2)'; Manufacturer = 'Port' }
  'easyrpg'        = @{ Platform = 'EasyRPG (RPG Maker)'; Manufacturer = 'Port' }
  'ecwolf'         = @{ Platform = 'Wolfenstein 3D'; Manufacturer = 'Port' }
  'eduke32'        = @{ Platform = 'Duke Nukem 3D'; Manufacturer = 'Port' }
  'electron'       = @{ Platform = 'Electron'; Manufacturer = 'Acorn Computers' }
  'enterprise'     = @{ Platform = 'Enterprise'; Manufacturer = 'Enterprise Computers' }
  'etlegacy'       = @{ Platform = 'ET Legacy (Enemy Territory: Quake Wars)'; Manufacturer = 'Port' }
  'fallout1-ce'    = @{ Platform = 'Fallout CE'; Manufacturer = 'Port' }
  'fallout2-ce'    = @{ Platform = 'Fallout2 CE'; Manufacturer = 'Port' }
  'fbneo'          = @{ Platform = 'FinalBurn Neo'; Manufacturer = 'Various' }
  'fds'            = @{ Platform = 'Family Computer Disk System/Famicom'; Manufacturer = 'Nintendo' }
  'flash'          = @{ Platform = 'Flashpoint (Adobe Flash)'; Manufacturer = 'Bluemaxima' }
  'flatpak'        = @{ Platform = 'Flatpak'; Manufacturer = 'Linux' }
  'fm7'            = @{ Platform = 'Fujitsu Micro 7'; Manufacturer = 'Fujitsu' }
  'fmtowns'        = @{ Platform = 'FM Towns/Towns Marty'; Manufacturer = 'Fujitsu' }
  'fpinball'       = @{ Platform = 'Future Pinball'; Manufacturer = 'Port' }
  'fury'           = @{ Platform = 'Ion Fury'; Manufacturer = 'Port' }
  'gamate'         = @{ Platform = 'Gamate/Super Boy/Super Child Prodigy'; Manufacturer = 'Bit Corporation' }
  'gameandwatch'   = @{ Platform = 'Game & Watch'; Manufacturer = 'Nintendo' }
  'gamecom'        = @{ Platform = 'Game.com'; Manufacturer = 'Tiger Electronics' }
  'gamecube'       = @{ Platform = 'GameCube'; Manufacturer = 'Nintendo' }
  'gamegear'       = @{ Platform = 'Game Gear'; Manufacturer = 'Sega' }
  'gamepock'       = @{ Platform = 'Game Pocket Computer'; Manufacturer = 'Epoch' }
  'gb'             = @{ Platform = 'Game Boy'; Manufacturer = 'Nintendo' }
  'gb2players'     = @{ Platform = 'Game Boy 2 Players'; Manufacturer = 'Nintendo' }
  'gba'            = @{ Platform = 'Game Boy Advance'; Manufacturer = 'Nintendo' }
  'gbc'            = @{ Platform = 'Game Boy Color'; Manufacturer = 'Nintendo' }
  'gbc2players'    = @{ Platform = 'Game Boy Color 2 Players'; Manufacturer = 'Nintendo' }
  'gmaster'        = @{ Platform = 'Game Master/Systema 2000/Super Game/Game Tronic'; Manufacturer = 'Hartung, et al.' }
  'gp32'           = @{ Platform = 'GP32'; Manufacturer = 'Game Park' }
  'gx4000'         = @{ Platform = 'Amstrad GX4000'; Manufacturer = 'Amstrad' }
  'gzdoom'         = @{ Platform = 'GZDoom (Boom/Chex Quest/Heretic/Hexen/Strife)'; Manufacturer = 'Port' }
  'hcl'            = @{ Platform = 'Hydra Castle Labyrinth'; Manufacturer = 'Port' }
  'hurrican'       = @{ Platform = 'Hurrican'; Manufacturer = 'Port' }
  'ikemen'         = @{ Platform = 'Ikemen Go'; Manufacturer = 'Port' }
  'intellivision'  = @{ Platform = 'Intellivision'; Manufacturer = 'Mattel' }
  'iortcw'         = @{ Platform = 'io Return to Castle Wolfenstein'; Manufacturer = 'Port' }
  'jaguar'         = @{ Platform = 'Atari Jaguar'; Manufacturer = 'Atari' }
  'jaguarcd'       = @{ Platform = 'Atari Jaguar CD'; Manufacturer = 'Atari' }
  'laser310'       = @{ Platform = 'Laser 310'; Manufacturer = 'Video Technology (VTech)' }
  'lcdgames'       = @{ Platform = 'Handheld LCD Games'; Manufacturer = 'Various' }
  'lindbergh'      = @{ Platform = 'Lindbergh'; Manufacturer = 'Sega' }
  'lowresnx'       = @{ Platform = 'Lowres NX'; Manufacturer = 'Timo Kloss' }
  'lutro'          = @{ Platform = 'Lutro'; Manufacturer = 'Port' }
  'lynx'           = @{ Platform = 'Atari Lynx'; Manufacturer = 'Atari' }
  'macintosh'      = @{ Platform = 'Macintosh 128K'; Manufacturer = 'Apple' }
  'mame'           = @{ Platform = 'Multiple Arcade Machine Emulator'; Manufacturer = 'Various' }
  'mame/model1'    = @{ Platform = 'Model 1'; Manufacturer = 'Sega' }
  'mastersystem'   = @{ Platform = 'Master System/Mark III'; Manufacturer = 'Sega' }
  'megaduck'       = @{ Platform = 'Mega Duck/Cougar Boy'; Manufacturer = 'Welback Holdings' }
  'megadrive'      = @{ Platform = 'Genesis/Mega Drive'; Manufacturer = 'Sega' }
  'model2'         = @{ Platform = 'Model 2'; Manufacturer = 'Sega' }
  'model3'         = @{ Platform = 'Model 3'; Manufacturer = 'Sega' }
  'moonlight'      = @{ Platform = 'Moonlight'; Manufacturer = 'Port' }
  'mrboom'         = @{ Platform = 'Mr. Boom'; Manufacturer = 'Port' }
  'msu-md'         = @{ Platform = 'MSU-MD'; Manufacturer = 'Sega' }
  'msx1'           = @{ Platform = 'MSX1'; Manufacturer = 'Microsoft' }
  'msx2'           = @{ Platform = 'MSX2'; Manufacturer = 'Microsoft' }
  'msx2+'          = @{ Platform = 'MSX2plus'; Manufacturer = 'Microsoft' }
  'msxturbor'      = @{ Platform = 'MSX TurboR'; Manufacturer = 'Microsoft' }
  'multivision'    = @{ Platform = 'Othello_Multivision'; Manufacturer = 'Tsukuda Original' }
  'mugen'          = @{ Platform = 'M.U.G.E.N'; Manufacturer = 'Port' }
  'n64'            = @{ Platform = 'Nintendo 64'; Manufacturer = 'Nintendo' }
  'n64dd'          = @{ Platform = 'Nintendo 64DD'; Manufacturer = 'Nintendo' }
  'namco2x6'       = @{ Platform = 'Namco System 246'; Manufacturer = 'Sony / Namco' }
  'naomi'          = @{ Platform = 'NAOMI'; Manufacturer = 'Sega' }
  'naomi2'         = @{ Platform = 'NAOMI 2'; Manufacturer = 'Sega' }
  'nds'            = @{ Platform = 'Nintendo DS'; Manufacturer = 'Nintendo' }
  'neogeo'         = @{ Platform = 'Neo Geo'; Manufacturer = 'SNK' }
  'neogeocd'       = @{ Platform = 'Neo Geo CD'; Manufacturer = 'SNK' }
  'nes'            = @{ Platform = 'Nintendo Entertainment System/Famicom'; Manufacturer = 'Nintendo' }
  'ngp'            = @{ Platform = 'Neo Geo Pocket'; Manufacturer = 'SNK' }
  'ngpc'           = @{ Platform = 'Neo Geo Pocket Color'; Manufacturer = 'SNK' }
  'o2em'           = @{ Platform = 'Odyssey 2/Videopac G7000'; Manufacturer = 'Magnavox/Philips' }
  'odcommander'    = @{ Platform = 'OD Commander'; Manufacturer = 'Port File Manager' }
  'odyssey2'       = @{ Platform = 'Odyssey 2/Videopac G7000'; Manufacturer = 'Magnavox/Philips' }
  'openbor'        = @{ Platform = 'Open Beats of Rage'; Manufacturer = 'Port' }
  'openjazz'       = @{ Platform = 'Openjazz'; Manufacturer = 'Port' }
  'openlara'       = @{ Platform = 'Tomb Raider'; Manufacturer = 'Port' }
  'oricatmos'      = @{ Platform = 'Oric Atmos'; Manufacturer = 'Tangerine Computer Systems' }
  'pc60'           = @{ Platform = 'NEC PC-6000'; Manufacturer = 'NEC' }
  'pc88'           = @{ Platform = 'NEC PC-8800'; Manufacturer = 'NEC' }
  'pc98'           = @{ Platform = 'NEC PC-9800/PC-98'; Manufacturer = 'NEC' }
  'pcengine'       = @{ Platform = 'PC Engine/TurboGrafx-16'; Manufacturer = 'NEC' }
  'pcenginecd'     = @{ Platform = 'PC Engine CD-ROM2/Duo R/Duo RX/TurboGrafx CD/TurboDuo'; Manufacturer = 'NEC' }
  'pcfx'           = @{ Platform = 'NEC PC-FX'; Manufacturer = 'NEC' }
  'pdp1'           = @{ Platform = 'PDP-1'; Manufacturer = 'Digital Equipment Corporation' }
  'pet'            = @{ Platform = 'Commodore PET'; Manufacturer = 'Commodore' }
  'pico'           = @{ Platform = 'Pico'; Manufacturer = 'Sega' }
  'pico8'          = @{ Platform = 'PICO-8 Fantasy Console'; Manufacturer = 'Lexaloffle Games' }
  'plugnplay'      = @{ Platform = 'Plug ''n'' Play/Handheld TV Games'; Manufacturer = 'Various' }
  'pokemini'       = @{ Platform = 'Pokemon Mini'; Manufacturer = 'Nintendo' }
  'ports'          = @{ Platform = 'Native ports'; Manufacturer = 'Linux' }
  'prboom'         = @{ Platform = 'Proff Boom'; Manufacturer = 'Port' }
  'ps2'            = @{ Platform = 'PlayStation 2'; Manufacturer = 'Sony' }
  'ps3'            = @{ Platform = 'PlayStation 3'; Manufacturer = 'Sony' }
  'ps4'            = @{ Platform = 'PlayStation 4'; Manufacturer = 'Sony' }
  'psp'            = @{ Platform = 'PlayStation Portable'; Manufacturer = 'Sony' }
  'psvita'         = @{ Platform = 'Vita'; Manufacturer = 'Sony' }
  'psx'            = @{ Platform = 'PlayStation'; Manufacturer = 'Sony' }
  'pv1000'         = @{ Platform = 'Casio PV-1000'; Manufacturer = 'Casio' }
  'pygame'         = @{ Platform = 'Python Games'; Manufacturer = 'Port' }
  'pyxel'          = @{ Platform = 'Pyxel Fantasy Console'; Manufacturer = 'Takashi Kitao' }
  'quake3'         = @{ Platform = 'Quake 3'; Manufacturer = 'Port' }
  'raze'           = @{ Platform = 'Raze'; Manufacturer = 'Port' }
  'reminiscence'   = @{ Platform = 'Reminiscence (Flashback Emulator)'; Manufacturer = 'Port' }
  'retroarch'      = @{ Platform = 'RetroArch (Liberato)'; Manufacturer = 'Hans-Kristian "Themaister" Arntzen' }
  'samcoupe'       = @{ Platform = 'SAM Coupe'; Manufacturer = 'Miles Gordon Technology' }
  'satellaview'    = @{ Platform = 'Satellaview'; Manufacturer = 'Nintendo' }
  'saturn'         = @{ Platform = 'Saturn'; Manufacturer = 'Sega' }
  'scummvm'        = @{ Platform = 'ScummVM'; Manufacturer = 'Ludvig Strigeus/Vincent Hamm' }
  'scv'            = @{ Platform = 'Super Cassette Vision'; Manufacturer = 'Epoch Co.' }
  'sdlpop'         = @{ Platform = 'SDLPoP (Prince of Persia)'; Manufacturer = 'Port' }
  'sega32x'        = @{ Platform = 'Sega 32X'; Manufacturer = 'Sega' }
  'segacd'         = @{ Platform = 'Sega CD/Mega CD'; Manufacturer = 'Sega' }
  'sg1000'         = @{ Platform = 'SG-1000/SG-1000 II/SC-3000'; Manufacturer = 'Sega' }
  'sgb'            = @{ Platform = 'Super Game Boy'; Manufacturer = 'Nintendo' }
  'sgb-msu1'       = @{ Platform = 'LADX-MSU1'; Manufacturer = 'Nintendo' }
  'singe'          = @{ Platform = 'SINGE'; Manufacturer = 'Various' }
  'snes'           = @{ Platform = 'Super Nintendo Entertainment System'; Manufacturer = 'Nintendo' }
  'snes-msu1'      = @{ Platform = 'Super NES CD-ROM/SNES MSU-1'; Manufacturer = 'Nintendo' }
  'socrates'       = @{ Platform = 'Socrates'; Manufacturer = 'VTech' }
  'solarus'        = @{ Platform = 'Solarus'; Manufacturer = 'Port' }
  'sonic-mania'    = @{ Platform = 'Sonic Mania'; Manufacturer = 'Port' }
  'sonic3-air'     = @{ Platform = 'Sonic 3 Angel Island Revisited'; Manufacturer = 'Port' }
  'sonicretro'     = @{ Platform = 'Star Engine/Sonic Retro Engine'; Manufacturer = 'Port' }
  'spectravideo'   = @{ Platform = 'Spectravideo'; Manufacturer = 'Spectravideo' }
  'steam'          = @{ Platform = 'Steam'; Manufacturer = 'Valve' }
  'sufami'         = @{ Platform = 'SuFami Turbo'; Manufacturer = 'Bandai' }
  'superbroswar'   = @{ Platform = 'Super Mario War'; Manufacturer = 'Port' }
  'supergrafx'     = @{ Platform = 'PC Engine/SuperGrafx/PC Engine 2'; Manufacturer = 'NEC' }
  'supervision'    = @{ Platform = 'Watara Supervision'; Manufacturer = 'Watara' }
  'supracan'       = @{ Platform = 'Super A''Can'; Manufacturer = 'Funtech Entertainment' }
  'switch'         = @{ Platform = 'Switch'; Manufacturer = 'Nintendo' }
  'systemsp'       = @{ Platform = 'Sega System SP'; Manufacturer = 'Sega' }
  'theforceengine' = @{ Platform = 'The Force Engine (Star Wars: Dark Forces)'; Manufacturer = 'Port' }
  'thextech'       = @{ Platform = 'TheXTech (Mega man)'; Manufacturer = 'Sinclair' }
  'thomson'        = @{ Platform = 'Thomson MO/TO Series Computer'; Manufacturer = 'Thomson' }
  'ti99'           = @{ Platform = 'TI-99/4/4A'; Manufacturer = 'Texas Instruments' }
  'tic80'          = @{ Platform = 'TIC-80 Fantasy Console'; Manufacturer = 'Vadim Grigoruk' }
  'traider1'       = @{ Platform = 'TR1X (Tomb Raider 1)'; Manufacturer = 'Port' }
  'traider2'       = @{ Platform = 'TR2X (Tomb Rauder 2)'; Manufacturer = 'Port' }
  'triforce'       = @{ Platform = 'Triforce'; Manufacturer = 'Namco/Sega/Nintendo' }
  'tutor'          = @{ Platform = 'Tomy Tutor/Pyuta/Grandstand Tutor'; Manufacturer = 'Tomy' }
  'tyrain'         = @{ Platform = 'TyrQuake (Quake)'; Manufacturer = 'Port' }
  'tyrquake'       = @{ Platform = 'TyrQuake (Quake 1)'; Manufacturer = 'Port' }
  'uqm'            = @{ Platform = 'The Ur-Quan Master (Star Control II)'; Manufacturer = 'Port' }
  'uzebox'         = @{ Platform = 'Uzebox Open-Source Console'; Manufacturer = 'Alec Bourque' }
  'vectrex'        = @{ Platform = 'Vectrex'; Manufacturer = 'Milton Bradley' }
  'vc4000'         = @{ Platform = 'Video Computer 4000'; Manufacturer = 'Interton' }
  'vgmplay'        = @{ Platform = 'MAME Video Game Music Player'; Manufacturer = 'Various' }
  'vircon32'       = @{ Platform = 'Vircon32 Virtual Console'; Manufacturer = 'Carra' }
  'vis'            = @{ Platform = 'Video Information System'; Manufacturer = 'Tandy/Memorex' }
  'vitaquake2'     = @{ Platform = 'PlayStation Vita port of Quake II'; Manufacturer = 'Port' }
  'virtualboy'     = @{ Platform = 'Virtual Boy'; Manufacturer = 'Nintendo' }
  'vpinball'       = @{ Platform = 'Visual Pinball'; Manufacturer = 'Port' }
  'voxatron'       = @{ Platform = 'Voxatron Fantasy Console'; Manufacturer = 'Lexaloffle Games' }
  'vsmile'         = @{ Platform = 'V.Smile (TV LEARNING SYSTEM)'; Manufacturer = 'VTech' }
  'wasm4'          = @{ Platform = 'WASM4 Fantasy Console'; Manufacturer = 'Aduros' }
  'wii'            = @{ Platform = 'Wii'; Manufacturer = 'Nintendo' }
  'wiiu'           = @{ Platform = 'Wii U'; Manufacturer = 'Nintendo' }
  'windows'        = @{ Platform = 'WINE'; Manufacturer = 'Bob Amstadt/Alexandre Julliard' }
  'wswan'          = @{ Platform = 'WonderSwan'; Manufacturer = 'Bandai' }
  'wswanc'         = @{ Platform = 'WonderSwan Color'; Manufacturer = 'Bandai' }
  'x1'             = @{ Platform = 'Sharp X1'; Manufacturer = 'Sharp' }
  'x68000'         = @{ Platform = 'Sharp X68000'; Manufacturer = 'Sharp' }
  'xash3d_fwgs'    = @{ Platform = 'Xash3D FWGS (Valve Games)'; Manufacturer = 'Port' }
  'xbox'           = @{ Platform = 'Xbox'; Manufacturer = 'Microsoft' }
  'xbox360'        = @{ Platform = 'Xbox 360'; Manufacturer = 'Microsoft' }
  'xegs'           = @{ Platform = 'Atari XEGS'; Manufacturer = 'Atari' }
  'xrick'          = @{ Platform = 'Rick Dangerous'; Manufacturer = 'Port' }
  'zx81'           = @{ Platform = 'Sinclair ZX81'; Manufacturer = 'Sinclair' }
  'zxspectrum'     = @{ Platform = 'ZX Spectrum'; Manufacturer = 'Sinclair' }
}

# ==================================================================================================
# FUNCTIONS
# ==================================================================================================

# --- FUNCTION: Get-PlatformInfo ---
# PURPOSE:
# - Translate a platform folder name (e.g., "psx") into a friendly platform name AND a Manufacturer.
# NOTES:
# - Falls back to returning the folder name as Platform and "" as Manufacturer if no translation is present.
function Get-PlatformInfo {
  param([string]$PlatformFolder)

  $k = ([string]$PlatformFolder).ToLowerInvariant()

  if ($PlatformMap.ContainsKey($k)) {
    $m = $PlatformMap[$k]
    return [pscustomobject]@{
      PlatformName = [string]$m.Platform
      Manufacturer = [string]$m.Manufacturer
    }
  }

  return [pscustomobject]@{
    PlatformName = [string]$PlatformFolder
    Manufacturer = ''
  }
}

# --- FUNCTION: Get-FriendlyPlatformName ---
# PURPOSE:
# - Preserve existing callers that only need a name string.
# NOTES:
# - Returns the PlatformName portion of Get-PlatformInfo.
function Get-FriendlyPlatformName {
  param([string]$PlatformFolder)
  return (Get-PlatformInfo -PlatformFolder $PlatformFolder).PlatformName
}

# --------------------------------------------------------------------------------------------------
# Mode banner
# PURPOSE:
# - Provide an immediate, unambiguous indication of what scope will be processed
# NOTES:
# - ROMS root platform count is printed after discovery is complete
# - Single-platform prints folder + friendly platform name immediately
# --------------------------------------------------------------------------------------------------
if ($isSinglePlatformMode) {
  $modePlatformFolder = Split-Path -Leaf $startDir
  $modePlatformName   = Get-FriendlyPlatformName $modePlatformFolder
  Write-Host ("MODE: Single-platform ({0} / {1})" -f $modePlatformFolder, $modePlatformName) -ForegroundColor Green
} else {
  Write-Host "MODE: ROMS root (discovering platforms...)" -ForegroundColor Green
}

# --- FUNCTION: Convert-ToTitleCaseSafe ---
# PURPOSE:
# - Convert a string into Title Case.
# NOTES:
# - Preserves tokens that look like acronyms/codes or contain digits.
function Convert-ToTitleCaseSafe {
  param([object]$Text)

  if ($null -eq $Text) { return '' }
  $t = ([string]$Text).Trim()
  if ([string]::IsNullOrWhiteSpace($t)) { return $t }

  $ti = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
  $tokens = $t -split '(\s+)'

  $out = foreach ($tok in $tokens) {
    if ($tok -match '^\s+$') { $tok; continue }
    if ($tok -match '\d' -or ($tok -cmatch '^[A-Z0-9&\-\+]{2,}$')) { $tok }
    else { $ti.ToTitleCase($tok.ToLowerInvariant()) }
  }

  return ($out -join '')
}

# --- FUNCTION: Get-DisplayNameFallback ---
# PURPOSE:
# - Provide a stable display name when <name> is missing/blank by falling back to the filename from <path>.
# NOTES:
# - Returns "" if neither a usable name nor a usable path is available.
function Get-DisplayNameFallback {
  param(
    [AllowNull()][AllowEmptyString()][string]$Name,
    [AllowNull()][AllowEmptyString()][string]$Path
  )

  $n = [string]$Name
  $p = [string]$Path

  if (-not [string]::IsNullOrWhiteSpace($n)) { return $n }

  if (-not [string]::IsNullOrWhiteSpace($p)) {
    try {
      return [System.IO.Path]::GetFileNameWithoutExtension($p)
    } catch {
      return $p
    }
  }

  return ""
}

# --- FUNCTION: Get-TitleForOutput ---
# PURPOSE:
# - Produce the CSV Title value with safe fallbacks for rare malformed/partial entries.
# NOTES:
# - Fallback ladder:
#     1) Title Case of the resolved display name
#     2) Raw resolved display name as-is
#     3) "(Untitled)" if nothing is available
function Get-TitleForOutput {
  param(
    [AllowNull()][AllowEmptyString()][string]$ResolvedName
  )

  $r = [string]$ResolvedName
  $tc = Convert-ToTitleCaseSafe -Text $r

  if (-not [string]::IsNullOrWhiteSpace($tc)) { return $tc }
  if (-not [string]::IsNullOrWhiteSpace($r))  { return $r }

  return "(Untitled)"
}

# --- FUNCTION: Get-PlatformTargets ---
# PURPOSE:
# - Determine which platform folder(s) to scan based on where the script resides.
# NOTES:
# - If the script directory contains gamelist.xml, treat it as a single-platform run.
# - Otherwise, treat the script directory as ROMS root and scan all first-level subfolders for gamelist.xml.
# - In ROMS root mode, explicitly ignores non-platform folders (e.g. windows_installers).
function Get-PlatformTargets {
  param([string]$StartDir)

  $start = (Resolve-Path -LiteralPath $StartDir).Path
  $localGamelist = Join-Path $start 'gamelist.xml'

  if (Test-Path -LiteralPath $localGamelist) {
    return @([pscustomobject]@{
      PlatformFolder = (Split-Path -Leaf $start)
      GamelistPath   = $localGamelist
    })
  }

  # Folder ignore list (ROMS root mode only)
  # PURPOSE:
  # - Prevent non-platform folders from being treated as platforms during discovery.
  # NOTES:
  # - Batocera ROMS roots often contain utility folders that can have deep structures but no gamelist.xml.
  # - Keep this list minimal and explicit to avoid hiding legitimate platform folders.
  $ignoredFolders = @(
    'windows_installers'
  )

  $targets = @()
  Get-ChildItem -LiteralPath $start -Directory | ForEach-Object {

    # Ignore any known non-platform folders early (case-insensitive compare)
    if ($ignoredFolders -contains $_.Name.ToLowerInvariant()) { return }

    $g = Join-Path $_.FullName 'gamelist.xml'
    if (Test-Path -LiteralPath $g) {
      $targets += [pscustomobject]@{
        PlatformFolder = $_.Name
        GamelistPath   = $g
      }
    }
  }

  return $targets
}

# --- FUNCTION: Get-XmlNodeText ---
# PURPOSE:
# - Safely read an XML child node's InnerText without throwing when missing.
# NOTES:
# - Returns "" when the requested child node does not exist.
# - Uses local-name() matching so default namespaces do not break child selection.
# - Always returns a string.
function Get-XmlNodeText {
  param(
    [Parameter(Mandatory=$true)][System.Xml.XmlNode]$Node,
    [Parameter(Mandatory=$true)][string]$ChildName
  )

  if ($null -eq $Node) { return '' }
  if ([string]::IsNullOrWhiteSpace($ChildName)) { return '' }

  $child = $null
  try {
    $child = $Node.SelectSingleNode("*[local-name()='$ChildName']")
  } catch {
    $child = $null
  }

  if ($null -eq $child) { return '' }

  $text = ''
  try { $text = [string]$child.InnerText } catch { $text = '' }
  return $text.Trim()
}

# --- FUNCTION: Get-M3UDiskCount ---
# PURPOSE:
# - Count the number of disk entries inside an .m3u playlist file.
# NOTES:
# - Counts non-empty lines that do not start with '#'
# - Returns 1 if the file cannot be read or contains no usable entries
function Get-M3UDiskCount {
  param([Parameter(Mandatory=$true)][string]$M3UPath)

  if ([string]::IsNullOrWhiteSpace($M3UPath)) { return 1 }
  if (-not (Test-Path -LiteralPath $M3UPath)) { return 1 }

  $lines = @()
  try {
    $lines = @(Get-Content -LiteralPath $M3UPath -ErrorAction Stop)
  } catch {
    return 1
  }

  $usable = @(
    $lines |
      ForEach-Object {
        if ($null -eq $_) { '' }
        else {
          $s = ([string]$_)
          $s = $s.TrimStart([char]0xFEFF)
          $s.Trim()
        }
      } |
      Where-Object { $_ -ne '' -and (-not $_.StartsWith('#')) }
  )

  if (@($usable).Count -lt 1) { return 1 }
  return [int]@($usable).Count
}

# --- FUNCTION: Read-Gamelist ---
# PURPOSE:
# - Read a gamelist.xml and return a flat list of entry objects with Name/Path/Hidden fields.
# NOTES:
# - StrictMode-safe:
#     - Missing <hidden> defaults to $false
#     - Missing <name> falls back to filename from <path> for display purposes
# - Normal XML parse:
#     - Uses XmlDocument + XPath and correctly enumerates each <game> node
# - Malformed gamelist handling:
#     - If the XML document is not well-formed, normal XML parsing will fail.
#     - In that case, this function falls back to a fragment parse of each <game> block.
# - GroupKey:
#     - Used for grouping entries into sets
#     - Primary: resolved name
#     - Fallback: path when name is missing/blank (prevents unrelated entries collapsing into one group)
# - XMLState:
#     - Returned objects include XMLState = Normal or Malformed so the CSV can show broken gamelist sources.
# - Always returns an array (even if 0 or 1 item).
function Read-Gamelist {
  param([string]$PlatformFolder, [string]$GamelistPath)

  # Read full XML as a single raw string so we can:
  # - Attempt a standard DOM parse first
  # - Fall back to fragment salvage parsing if the document is malformed
  $raw = Get-Content -LiteralPath $GamelistPath -Raw
  if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

  $out = @()

  $doc = $null
  $parsedOk = $false
  try {
    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.LoadXml($raw)
    $parsedOk = $true
  } catch {
    $parsedOk = $false
  }

  if ($parsedOk -and $null -ne $doc) {

    # Normal parse path:
    # - Use XPath with local-name() so namespaces don't break discovery
    $nodes = $null
    try {
      $nodes = $doc.SelectNodes("//*[local-name()='game']")
    } catch {
      $nodes = $null
    }

    if ($null -ne $nodes) {
      foreach ($node in $nodes) {

        $name   = Get-XmlNodeText -Node $node -ChildName 'name'
        $path   = Get-XmlNodeText -Node $node -ChildName 'path'
        $hidden = Get-XmlNodeText -Node $node -ChildName 'hidden'

        $resolvedName = Get-DisplayNameFallback -Name $name -Path $path

        # GroupKey determines multi-disk grouping:
        # - Prefer the raw <name> when present (stable grouping)
        # - Fall back to path when name is missing to avoid unrelated collisions
        $groupKey     = if (-not [string]::IsNullOrWhiteSpace($name)) { $name.Trim() } else { [string]$path }

        $out += [pscustomobject]@{
          PlatformFolder = $PlatformFolder
          NameRaw        = [string]$name
          NameResolved   = [string]$resolvedName
          PathRaw        = [string]$path
          Hidden         = ($hidden -match '^(true|1|yes)$')
          XMLState       = 'Normal'
          GroupKey       = [string]$groupKey
        }
      }
    }

    return $out
  }

  # Salvage mode:
  # - Extract <game>...</game> blocks with regex
  # - Parse each block as a fragment XML document
  # - This allows recovery even when the overall gamelist.xml is malformed
  $gameBlocks = [regex]::Matches($raw, '(?is)<game\b[^>]*>.*?</game>')

  foreach ($m in $gameBlocks) {

    $block = $m.Value
    if ([string]::IsNullOrWhiteSpace($block)) { continue }

    $fragDoc = $null
    try {
      $fragDoc = New-Object System.Xml.XmlDocument
      $fragDoc.XmlResolver = $null
      $fragDoc.LoadXml("<root>$block</root>")
    } catch {
      continue
    }

    $node = $null
    try {
      $node = $fragDoc.SelectSingleNode("//*[local-name()='game']")
    } catch {
      $node = $null
    }

    if ($null -eq $node) { continue }

    $name   = Get-XmlNodeText -Node $node -ChildName 'name'
    $path   = Get-XmlNodeText -Node $node -ChildName 'path'
    $hidden = Get-XmlNodeText -Node $node -ChildName 'hidden'

    $resolvedName = Get-DisplayNameFallback -Name $name -Path $path
    $groupKey     = if (-not [string]::IsNullOrWhiteSpace($name)) { $name.Trim() } else { [string]$path }

    $out += [pscustomobject]@{
      PlatformFolder = $PlatformFolder
      NameRaw        = [string]$name
      NameResolved   = [string]$resolvedName
      PathRaw        = [string]$path
      Hidden         = ($hidden -match '^(true|1|yes)$')
      XMLState       = 'Malformed'
      GroupKey       = [string]$groupKey
    }
  }

  return $out
}

# --- FUNCTION: Write-CsvUtf8NoBom ---
# PURPOSE:
# - Write a comma-delimited CSV as UTF-8 without BOM for reliable Excel opening
# NOTES:
# - Avoids UTF-8 BOM characters appearing as ï»¿ when the file is interpreted incorrectly
# - Uses ConvertTo-Csv to guarantee comma delimiter and consistent quoting
function Write-CsvUtf8NoBom {
  param(
    [Parameter(Mandatory=$true)]$Rows,
    [Parameter(Mandatory=$true)][string]$Path
  )

  # ConvertTo-Csv produces a standards-compliant CSV with consistent quoting and comma separators.
  # We then write it as UTF-8 without BOM for maximum compatibility.
  $csvLines = @($Rows | ConvertTo-Csv -NoTypeInformation)

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllLines($Path, $csvLines, $utf8NoBom)
}

# ==================================================================================================
# PHASE 1: TARGET DISCOVERY
# ==================================================================================================

Write-Phase "Discovering platform folders and locating gamelist.xml files..."

$targets = @(Get-PlatformTargets -StartDir $startDir)

if (@($targets).Count -eq 0) {
  Write-Warning "No gamelist.xml found. Run from /roms or a platform folder that contains gamelist.xml."
  Write-RuntimeReport -Stop
  return
}

Write-Host "Found $($targets.Count) platform(s) with gamelist.xml"

if ($isRomsRootMode) {
  Write-Host ("MODE: ROMS root ({0} platform(s))" -f $targets.Count) -ForegroundColor Green
}

# ==================================================================================================
# PHASE 2: READ + GROUP ENTRIES / BUILD OUTPUT ROWS
# ==================================================================================================

Write-Phase "Reading gamelist.xml and collecting game entries..."

$rows = @()

foreach ($t in $targets) {

  if ($isRomsRootMode) {
    Write-Host ""
    $pf = [string]$t.PlatformFolder
    $pi = Get-PlatformInfo $pf
    Write-Host ("{0} | {1}" -f $pf, $pi.PlatformName) -ForegroundColor Green
  }

  $platformFolder         = [string]$t.PlatformFolder
  $platformInfo           = Get-PlatformInfo $platformFolder
  $platformName           = [string]$platformInfo.PlatformName
  $platformManufacturer   = [string]$platformInfo.Manufacturer

  # Read and normalize all <game> entries for the platform (normal parse or salvage mode).
  $entries = @(Read-Gamelist $platformFolder $t.GamelistPath)
  if ($entries.Count -eq 0) {
    if ($isRomsRootMode) {
      Write-Host "No entries found (skipping)." -ForegroundColor Yellow
      Write-Host "Finished platform: $($t.PlatformFolder)" -ForegroundColor Cyan
    }
    continue
  } else {
    if ($isRomsRootMode) {
      $entryLabel = if ($entries.Count -eq 1) { "entry" } else { "entries" }
      Write-Host ("{0} {1} found." -f $entries.Count, $entryLabel) -ForegroundColor DarkYellow
    }
  }

  # Group entries into "sets" so multi-disk collections can be inferred from:
  # - .m3u visible entries (Multi-M3U)
  # - multiple entries sharing the group key where additional disks are hidden (Multi-XML)
  foreach ($group in @($entries | Group-Object GroupKey)) {

    $items        = @($group.Group)
    $hiddenItems  = @($items | Where-Object { $_.Hidden })
    $visibleItems = @($items | Where-Object { -not $_.Hidden })

    # - If every entry in a group is hidden, still emit one row (otherwise the group disappears from the report).
    if ($visibleItems.Count -eq 0 -and $items.Count -gt 0) {
      $visibleItems = @($items | Select-Object -First 1)
    }

    # One row per set:
    # - If multiple entries are visible in the same group, pick a stable representative
    #   (otherwise the set would emit multiple rows).
    if ($visibleItems.Count -gt 1) {
      $visibleItems = @($visibleItems | Sort-Object PathRaw | Select-Object -First 1)
    }

    foreach ($g in $visibleItems) {

      $pathStr  = [string]$g.PathRaw
      $nameStr  = [string]$g.NameResolved
      $nameRaw  = [string]$g.NameRaw

      $isNotGame = $false
      if (-not [string]::IsNullOrWhiteSpace($nameRaw)) {
        $isNotGame = ($nameRaw -match '(?i)^\s*ZZZ\(notgame\):')
      }

      $typeTag = if ($isNotGame) { 'App' } else { 'Game' }

      $titleName = $nameStr
      if ($isNotGame) {
        $titleName = ($titleName -replace '(?i)^\s*ZZZ\(notgame\):\s*', '')
      }

      # Determine multi-disk behavior based on:
      # - .m3u path (Multi-M3U)
      # - presence of hidden sibling entries within the group (Multi-XML)
      $entryType = 'Single'
      if ($pathStr -match '(?i)\.m3u$') {
        $entryType = 'Multi-M3U'
      }
      else {
        # Multi-XML should only mean "this visible entry has hidden sibling entries in the same group".
        # This prevents single hidden entries (or all-hidden groups) from being mislabeled as Multi-XML.
        $hasHidden = ($hiddenItems.Count -gt 0)
        $hasVisible = ($items.Count -gt $hiddenItems.Count)  # at least one non-hidden exists in group
        $hasMultiple = ($items.Count -gt 1)

        if ($hasMultiple -and $hasHidden -and $hasVisible) {
          $entryType = 'Multi-XML'
        }
      }

      $title = Get-TitleForOutput -ResolvedName $titleName

      # DiskCount handling is based on the inferred entry type:
      # - Multi-M3U: read the .m3u file and count usable entries
      # - Multi-XML: count group entries (visible + hidden)
      # - Single: 1
      $diskCount = 1

      if ($entryType -eq 'Multi-M3U') {

        # Resolve .m3u path relative to the platform root folder when the gamelist uses ./ style paths.
        $platformRoot = $startDir
        if ($isRomsRootMode) {
          $platformRoot = Join-Path $startDir $platformFolder
        }

        $m3uFullPath = $pathStr
        if (-not [System.IO.Path]::IsPathRooted($m3uFullPath)) {
          $rel = $m3uFullPath.Trim()
          if ($rel.StartsWith('./')) { $rel = $rel.Substring(2) }
          elseif ($rel.StartsWith('.\')) { $rel = $rel.Substring(2) }
          $m3uFullPath = Join-Path $platformRoot $rel
        }

        try {
          $m3uFullPath = (Resolve-Path -LiteralPath $m3uFullPath -ErrorAction Stop).Path
        } catch {
          # Leave as-is; Get-M3UDiskCount will safely return 1 if unreadable/unresolvable.
        }

        $diskCount = Get-M3UDiskCount -M3UPath $m3uFullPath
      }
      elseif ($entryType -eq 'Multi-XML') {
        $diskCount = [int]$items.Count
      }
      else {
        $diskCount = 1
      }

      # Build output row in the exact column order expected for the CSV export
      $hiddenMark = if ($g.Hidden) { 'X' } else { '' }

      $rows += [pscustomobject]@{
        Title          = $title
        PlatformName   = $platformName
        Manufacturer   = $platformManufacturer
        Type           = [string]$typeTag
        DiscType       = [string]$entryType
        DiskCount      = [int]$diskCount
        Hidden         = [string]$hiddenMark
        PlatformFolder = $platformFolder
        FilePath       = $pathStr
        XMLState       = [string]$g.XMLState
      }
    }
  }

  if ($isRomsRootMode) {
    Write-Host "Finished platform: $($t.PlatformFolder)" -ForegroundColor Cyan
  }
}

Write-Phase "Finished collecting game entries."

# ==================================================================================================
# PHASE 3: SORT + EXPORT CSV
# ==================================================================================================

Write-Phase "Sorting and exporting CSV..."

# Sort for human browsing:
# - PlatformName groups platforms logically
# - Title alphabetizes games within a platform
$rows = @($rows | Sort-Object PlatformName, Title)

$outPath = Join-Path $startDir 'Game List.csv'

# --------------------------------------------------------------------------------------------------
# Output file handling
# PURPOSE:
# - Ensure any prior output file is removed so the generated CSV is a clean overwrite
# NOTES:
# - If the output file is locked (e.g., open in Excel), print a friendly message and abort cleanly
# --------------------------------------------------------------------------------------------------
if (Test-Path -LiteralPath $outPath) {
  try {
    Remove-Item -LiteralPath $outPath -Force -ErrorAction Stop
  } catch {
    Write-Host ""
    Write-Host "Cannot write 'Game List.csv' because the existing file is locked (likely open in Excel)." -ForegroundColor Yellow
    Write-Host "Script aborted." -ForegroundColor Yellow
    Write-RuntimeReport -Stop
    return
  }
}

Write-CsvUtf8NoBom -Rows $rows -Path $outPath

# --------------------------------------------------------------------------------------------------
# Malformed gamelist warning and gamelist path linkage
# PURPOSE:
# - Alert the user when one or more platforms required malformed XML handling
# - Include the gamelist.xml path for quick remediation
# NOTES:
# - This does not affect CSV output; it is console diagnostics only
# --------------------------------------------------------------------------------------------------
$malformedPlatforms = @(
  $rows |
    Where-Object { $_.XMLState -eq 'Malformed' } |
    Select-Object -ExpandProperty PlatformFolder -Unique |
    Sort-Object
)

if ($malformedPlatforms.Count -gt 0) {

  $gamelistPathByPlatform = @{}
  foreach ($t in $targets) {
    if ($null -ne $t -and $null -ne $t.PlatformFolder -and $null -ne $t.GamelistPath) {
      $gamelistPathByPlatform[[string]$t.PlatformFolder] = [string]$t.GamelistPath
    }
  }

  Write-Host ""
  Write-Host "WARNING: Malformed gamelist.xml detected for the following platform(s):" -ForegroundColor Yellow

  foreach ($p in $malformedPlatforms) {
    if ($gamelistPathByPlatform.ContainsKey([string]$p)) {
      Write-Host ("  - {0} ({1})" -f $p, $gamelistPathByPlatform[[string]$p]) -ForegroundColor Yellow
    } else {
      Write-Host "  - $p" -ForegroundColor Yellow
    }
  }

  Write-Host ""
  Write-Host "Game entries were recovered successfully, but the source files should be" -ForegroundColor Yellow
  Write-Host "regenerated or repaired (e.g. via a gamelist update/scrape, manual edit, etc.)." -ForegroundColor Yellow
}

# --------------------------------------------------------------------------------------------------
# Final summary output
# PURPOSE:
# - Provide a clear completion line and key output facts
# NOTES:
# - @($rows).Count forces array semantics under StrictMode even when only one row exists
# --------------------------------------------------------------------------------------------------
Write-Phase "Finished."
Write-Host "CSV written to: $outPath" -ForegroundColor Green
Write-Host "Total games exported: $(@($rows).Count)" -ForegroundColor Green
Write-RuntimeReport -Stop
