<#
PURPOSE: Export a master list of games across Batocera platform folders by reading each platform's gamelist.xml
VERSION: 1.0
AUTHOR: Devin Kelley, Distant Thunderworks LLC

NOTES:
- Place this file into the ROMS folder to process all platforms, or in a platform's individual subfolder to process just that one.
- This script does NOT modify any files; it only reads gamelist.xml and produces a CSV report.
- Output CSV is written to the same directory where the PS1 script resides.
- If "Game List.csv" already exists in the output folder, it is deleted and replaced with a newly generated file.
- Multi-disk detection is inferred as:
    - Multi-M3U: the visible entry’s <path> ends in .m3u
    - Multi-XML: the visible entry has 1+ additional entries with the same group key where <hidden>true</hidden> is set
        - This is tailored to be run after first running the Generate Batocera Playlists.ps1 script
    - Single: neither of the above
- Robustness:
    - Some gamelist.xml files in the wild can be malformed (mismatched tags, partial writes, etc.).
    - This script will attempt a normal XML parse first, and if that fails it will fall back to a "salvage mode"
      that extracts <game>...</game> blocks and parses them individually.
- XMLState column:
    - Normal: gamelist.xml parsed cleanly as a complete XML document
    - Malformed: gamelist.xml was malformed; entries were extracted by parsing <game> fragments
- Progress / phase output:
    - Prints only major phase steps
    - If running from ROMS root (multi-platform mode), prints per-platform start + finished lines
    - Always prints a final "finished" summary

BREAKDOWN
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
    - Note (supplemental platform note derived from the platform mapping)
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

# Folder -> Platform + Note mapping (extend as needed)
# - Platform: friendly platform name (no parentheses)
# - Note: supplemental note (manufacturer, port, etc.) without parentheses; may be ""
$PlatformMap = @{
  '3do'            = @{ Platform = '3DO'; Note = 'Panasonic' }
  '3ds'            = @{ Platform = 'Nintendo 3DS'; Note = 'Nintendo' }
  'abuse'          = @{ Platform = 'Abuse SDL'; Note = 'Port' }
  'adam'           = @{ Platform = 'Coleco Adam'; Note = 'Coleco' }
  'advision'       = @{ Platform = 'Adventure Vision'; Note = 'Entex' }
  'amiga1200'      = @{ Platform = 'Amiga 1200/AGA'; Note = 'Commodore' }
  'amiga500'       = @{ Platform = 'Amiga 500/OCS/ECS'; Note = 'Commodore' }
  'amigacd32'      = @{ Platform = 'Amiga CD32'; Note = 'Commodore' }
  'amigacdtv'      = @{ Platform = 'Commodore CDTV'; Note = 'Commodore' }
  'amstradcpc'     = @{ Platform = 'Amstrad CPC'; Note = 'Amstrad' }
  'apfm1000'       = @{ Platform = 'APF-MP1000/MP-1000/M-1000'; Note = 'APF Electronics Inc.' }
  'apple2'         = @{ Platform = 'Apple II'; Note = 'Apple' }
  'apple2gs'       = @{ Platform = 'Apple IIGS'; Note = 'Apple' }
  'arcadia'        = @{ Platform = 'Arcadia 2001'; Note = 'Emerson Radio' }
  'archimedes'     = @{ Platform = 'Archimedes'; Note = 'Acorn Computers' }
  'arduboy'        = @{ Platform = 'Arduboy'; Note = 'Arduboy' }
  'astrocde'       = @{ Platform = 'Astrocade'; Note = 'Bally/Midway' }
  'atari2600'      = @{ Platform = 'Atari 2600/VCS'; Note = 'Atari' }
  'atari5200'      = @{ Platform = 'Atari 5200'; Note = 'Atari' }
  'atari7800'      = @{ Platform = 'Atari 7800'; Note = 'Atari' }
  'atari800'       = @{ Platform = 'Atari 800'; Note = 'Atari' }
  'atarist'        = @{ Platform = 'Atari ST'; Note = 'Atari' }
  'atom'           = @{ Platform = 'Atom'; Note = 'Acorn Computers' }
  'atomiswave'     = @{ Platform = 'Sammy Atomiswave'; Note = 'Sammy' }
  'bbc'            = @{ Platform = 'BBC Micro/Master/Archimedes'; Note = 'Acorn Computers' }
  'bennugd'        = @{ Platform = 'BennuGD'; Note = 'Game Development Suite' }
  'boom3'          = @{ Platform = 'Doom 3'; Note = 'Port' }
  'camplynx'       = @{ Platform = 'Camputers Lynx'; Note = 'Camputers' }
  'cannonball'     = @{ Platform = 'Cannonball'; Note = 'Port' }
  'casloopy'       = @{ Platform = 'Casio Loopy'; Note = 'Casio' }
  'catacombgl'     = @{ Platform = 'Catacomb GL'; Note = 'Port' }
  'cavestory'      = @{ Platform = 'Cave Story'; Note = 'Port' }
  'c128'           = @{ Platform = 'Commodore 128'; Note = 'Commodore' }
  'c20'            = @{ Platform = 'Commodore VIC-20/VC-20'; Note = 'Commodore' }
  'c64'            = @{ Platform = 'Commodore 64'; Note = 'Commodore' }
  'cdi'            = @{ Platform = 'Compact Disc Interactive/CD-i'; Note = 'Philips, et al.' }
  'cdogs'          = @{ Platform = 'C-Dogs'; Note = 'Port' }
  'cgenius'        = @{ Platform = 'Commander Genius (Commander Keen and Cosmos the Cosmic Adventure)'; Note = 'Port' }
  'channelf'       = @{ Platform = 'Fairchild Channel F'; Note = 'Fairchild' }
  'chihiro'        = @{ Platform = 'Chihiro'; Note = 'Sega' }
  'coco'           = @{ Platform = 'TRS-80/Color Computer'; Note = 'Tandy/RadioShack' }
  'colecovision'   = @{ Platform = 'ColecoVision'; Note = 'Coleco' }
  'commanderx16'   = @{ Platform = 'Commander X16'; Note = 'David Murray' }
  'corsixth'       = @{ Platform = 'CorsixTH (Theme Hospital)'; Note = 'Port' }
  'cplus4'         = @{ Platform = 'Commodore Plus/4'; Note = 'Commodore' }
  'crvision'       = @{ Platform = 'CreatiVision/Educat 2002/Dick Smith Wizzard/FunVision'; Note = 'VTech' }
  'daphne'         = @{ Platform = 'DAPHNE Laserdisc'; Note = 'Various' }
  'devilutionx'    = @{ Platform = 'DevilutionX (Diablo/Hellfire)'; Note = 'Port' }
  'dice'           = @{ Platform = 'Discrete Integrated Circuit Emulator'; Note = 'Various' }
  'dolphin'        = @{ Platform = 'Dolphin'; Note = 'GameCube/Wii Emulator' }
  'dos'            = @{ Platform = 'DOSbox'; Note = 'Peter Veenstra/Sjoerd van der Berg' }
  'dreamcast'      = @{ Platform = 'Dreamcast'; Note = 'Sega' }
  'dxx-rebirth'    = @{ Platform = 'DXX Rebirth (Descent/Descent 2)'; Note = 'Port' }
  'easyrpg'        = @{ Platform = 'EasyRPG (RPG Maker)'; Note = 'Port' }
  'ecwolf'         = @{ Platform = 'Wolfenstein 3D'; Note = 'Port' }
  'eduke32'        = @{ Platform = 'Duke Nukem 3D'; Note = 'Port' }
  'electron'       = @{ Platform = 'Electron'; Note = 'Acorn Computers' }
  'enterprise'     = @{ Platform = 'Enterprise'; Note = 'Enterprise Computers' }
  'etlegacy'       = @{ Platform = 'ET Legacy (Enemy Territory: Quake Wars)'; Note = 'Port' }
  'fallout1-ce'    = @{ Platform = 'Fallout CE'; Note = 'Port' }
  'fallout2-ce'    = @{ Platform = 'Fallout2 CE'; Note = 'Port' }
  'fbneo'          = @{ Platform = 'FinalBurn Neo'; Note = 'Various' }
  'fds'            = @{ Platform = 'Family Computer Disk System/Famicom'; Note = 'Nintendo' }
  'flash'          = @{ Platform = 'Flashpoint (Adobe Flash)'; Note = 'Bluemaxima' }
  'flatpak'        = @{ Platform = 'Flatpak'; Note = 'Linux' }
  'fm7'            = @{ Platform = 'Fujitsu Micro 7'; Note = 'Fujitsu' }
  'fmtowns'        = @{ Platform = 'FM Towns/Towns Marty'; Note = 'Fujitsu' }
  'fpinball'       = @{ Platform = 'Future Pinball'; Note = 'Port' }
  'fury'           = @{ Platform = 'Ion Fury'; Note = 'Port' }
  'gamate'         = @{ Platform = 'Gamate/Super Boy/Super Child Prodigy'; Note = 'Bit Corporation' }
  'gameandwatch'   = @{ Platform = 'Game & Watch'; Note = 'Nintendo' }
  'gamecom'        = @{ Platform = 'Game.com'; Note = 'Tiger Electronics' }
  'gamecube'       = @{ Platform = 'GameCube'; Note = 'Nintendo' }
  'gamegear'       = @{ Platform = 'Game Gear'; Note = 'Sega' }
  'gamepock'       = @{ Platform = 'Game Pocket Computer'; Note = 'Epoch' }
  'gb'             = @{ Platform = 'Game Boy'; Note = 'Nintendo' }
  'gb2players'     = @{ Platform = 'Game Boy 2 Players'; Note = 'Nintendo' }
  'gba'            = @{ Platform = 'Game Boy Advance'; Note = 'Nintendo' }
  'gbc'            = @{ Platform = 'Game Boy Color'; Note = 'Nintendo' }
  'gbc2players'    = @{ Platform = 'Game Boy Color 2 Players'; Note = 'Nintendo' }
  'gmaster'        = @{ Platform = 'Game Master/Systema 2000/Super Game/Game Tronic'; Note = 'Hartung, et al.' }
  'gp32'           = @{ Platform = 'GP32'; Note = 'Game Park' }
  'gx4000'         = @{ Platform = 'Amstrad GX4000'; Note = 'Amstrad' }
  'gzdoom'         = @{ Platform = 'GZDoom (Boom/Chex Quest/Heretic/Hexen/Strife)'; Note = 'Port' }
  'hcl'            = @{ Platform = 'Hydra Castle Labyrinth'; Note = 'Port' }
  'hurrican'       = @{ Platform = 'Hurrican'; Note = 'Port' }
  'ikemen'         = @{ Platform = 'Ikemen Go'; Note = 'Port' }
  'intellivision'  = @{ Platform = 'Intellivision'; Note = 'Mattel' }
  'iortcw'         = @{ Platform = 'io Return to Castle Wolfenstein'; Note = 'Port' }
  'jaguar'         = @{ Platform = 'Atari Jaguar'; Note = 'Atari' }
  'jaguarcd'       = @{ Platform = 'Atari Jaguar CD'; Note = 'Atari' }
  'laser310'       = @{ Platform = 'Laser 310'; Note = 'Video Technology (VTech)' }
  'lcdgames'       = @{ Platform = 'Handheld LCD Games'; Note = 'Various' }
  'lindbergh'      = @{ Platform = 'Lindbergh'; Note = 'Sega' }
  'lowresnx'       = @{ Platform = 'Lowres NX'; Note = 'Timo Kloss' }
  'lutro'          = @{ Platform = 'Lutro'; Note = 'Port' }
  'lynx'           = @{ Platform = 'Atari Lynx'; Note = 'Atari' }
  'macintosh'      = @{ Platform = 'Macintosh 128K'; Note = 'Apple' }
  'mame'           = @{ Platform = 'Multiple Arcade Machine Emulator'; Note = 'Various' }
  'mame/model1'    = @{ Platform = 'Model 1'; Note = 'Sega' }
  'mastersystem'   = @{ Platform = 'Master System/Mark III'; Note = 'Sega' }
  'megaduck'       = @{ Platform = 'Mega Duck/Cougar Boy'; Note = 'Welback Holdings' }
  'megadrive'      = @{ Platform = 'Genesis/Mega Drive'; Note = 'Sega' }
  'model2'         = @{ Platform = 'Model 2'; Note = 'Sega' }
  'model3'         = @{ Platform = 'Model 3'; Note = 'Sega' }
  'moonlight'      = @{ Platform = 'Moonlight'; Note = 'Port' }
  'mrboom'         = @{ Platform = 'Mr. Boom'; Note = 'Port' }
  'msu-md'         = @{ Platform = 'MSU-MD'; Note = 'Sega' }
  'msx1'           = @{ Platform = 'MSX1'; Note = 'Microsoft' }
  'msx2'           = @{ Platform = 'MSX2'; Note = 'Microsoft' }
  'msx2+'          = @{ Platform = 'MSX2plus'; Note = 'Microsoft' }
  'msxturbor'      = @{ Platform = 'MSX TurboR'; Note = 'Microsoft' }
  'multivision'    = @{ Platform = 'Othello_Multivision'; Note = 'Tsukuda Original' }
  'mugen'          = @{ Platform = 'M.U.G.E.N'; Note = 'Port' }
  'n64'            = @{ Platform = 'Nintendo 64'; Note = 'Nintendo' }
  'n64dd'          = @{ Platform = 'Nintendo 64DD'; Note = 'Nintendo' }
  'namco2x6'       = @{ Platform = 'Namco System 246'; Note = 'Sony / Namco' }
  'naomi'          = @{ Platform = 'NAOMI'; Note = 'Sega' }
  'naomi2'         = @{ Platform = 'NAOMI 2'; Note = 'Sega' }
  'nds'            = @{ Platform = 'Nintendo DS'; Note = 'Nintendo' }
  'neogeo'         = @{ Platform = 'Neo Geo'; Note = 'SNK' }
  'neogeocd'       = @{ Platform = 'Neo Geo CD'; Note = 'SNK' }
  'nes'            = @{ Platform = 'Nintendo Entertainment System/Famicom'; Note = 'Nintendo' }
  'ngp'            = @{ Platform = 'Neo Geo Pocket'; Note = 'SNK' }
  'ngpc'           = @{ Platform = 'Neo Geo Pocket Color'; Note = 'SNK' }
  'o2em'           = @{ Platform = 'Odyssey 2/Videopac G7000'; Note = 'Magnavox/Philips' }
  'odcommander'    = @{ Platform = 'OD Commander'; Note = 'Port File Manager' }
  'odyssey2'       = @{ Platform = 'Odyssey 2/Videopac G7000'; Note = 'Magnavox/Philips' }
  'openbor'        = @{ Platform = 'Open Beats of Rage'; Note = 'Port' }
  'openjazz'       = @{ Platform = 'Openjazz'; Note = 'Port' }
  'openlara'       = @{ Platform = 'Tomb Raider'; Note = 'Port' }
  'oricatmos'      = @{ Platform = 'Oric Atmos'; Note = 'Tangerine Computer Systems' }
  'pc60'           = @{ Platform = 'NEC PC-6000'; Note = 'NEC' }
  'pc88'           = @{ Platform = 'NEC PC-8800'; Note = 'NEC' }
  'pc98'           = @{ Platform = 'NEC PC-9800/PC-98'; Note = 'NEC' }
  'pcengine'       = @{ Platform = 'PC Engine/TurboGrafx-16'; Note = 'NEC' }
  'pcenginecd'     = @{ Platform = 'PC Engine CD-ROM2/Duo R/Duo RX/TurboGrafx CD/TurboDuo'; Note = 'NEC' }
  'pcfx'           = @{ Platform = 'NEC PC-FX'; Note = 'NEC' }
  'pdp1'           = @{ Platform = 'PDP-1'; Note = 'Digital Equipment Corporation' }
  'pet'            = @{ Platform = 'Commodore PET'; Note = 'Commodore' }
  'pico'           = @{ Platform = 'Pico'; Note = 'Sega' }
  'pico8'          = @{ Platform = 'PICO-8 Fantasy Console'; Note = 'Lexaloffle Games' }
  'plugnplay'      = @{ Platform = 'Plug ''n'' Play/Handheld TV Games'; Note = 'Various' }
  'pokemini'       = @{ Platform = 'Pokemon Mini'; Note = 'Nintendo' }
  'ports'          = @{ Platform = 'Native ports'; Note = 'Linux' }
  'prboom'         = @{ Platform = 'Proff Boom'; Note = 'Port' }
  'ps2'            = @{ Platform = 'PlayStation 2'; Note = 'Sony' }
  'ps3'            = @{ Platform = 'PlayStation 3'; Note = 'Sony' }
  'ps4'            = @{ Platform = 'PlayStation 4'; Note = 'Sony' }
  'psp'            = @{ Platform = 'PlayStation Portable'; Note = 'Sony' }
  'psvita'         = @{ Platform = 'Vita'; Note = 'Sony' }
  'psx'            = @{ Platform = 'PlayStation'; Note = 'Sony' }
  'pv1000'         = @{ Platform = 'Casio PV-1000'; Note = 'Casio' }
  'pygame'         = @{ Platform = 'Python Games'; Note = 'Port' }
  'pyxel'          = @{ Platform = 'Pyxel Fantasy Console'; Note = 'Takashi Kitao' }
  'quake3'         = @{ Platform = 'Quake 3'; Note = 'Port' }
  'raze'           = @{ Platform = 'Raze'; Note = 'Port' }
  'reminiscence'   = @{ Platform = 'Reminiscence (Flashback Emulator)'; Note = 'Port' }
  'retroarch'      = @{ Platform = 'RetroArch (Liberato)'; Note = 'Hans-Kristian "Themaister" Arntzen' }
  'samcoupe'       = @{ Platform = 'SAM Coupe'; Note = 'Miles Gordon Technology' }
  'satellaview'    = @{ Platform = 'Satellaview'; Note = 'Nintendo' }
  'saturn'         = @{ Platform = 'Saturn'; Note = 'Sega' }
  'scummvm'        = @{ Platform = 'ScummVM'; Note = 'Ludvig Strigeus/Vincent Hamm' }
  'scv'            = @{ Platform = 'Super Cassette Vision'; Note = 'Epoch Co.' }
  'sdlpop'         = @{ Platform = 'SDLPoP (Prince of Persia)'; Note = 'Port' }
  'sega32x'        = @{ Platform = 'Sega 32X'; Note = 'Sega' }
  'segacd'         = @{ Platform = 'Sega CD/Mega CD'; Note = 'Sega' }
  'sg1000'         = @{ Platform = 'SG-1000/SG-1000 II/SC-3000'; Note = 'Sega' }
  'sgb'            = @{ Platform = 'Super Game Boy'; Note = 'Nintendo' }
  'sgb-msu1'       = @{ Platform = 'LADX-MSU1'; Note = 'Nintendo' }
  'singe'          = @{ Platform = 'SINGE'; Note = 'Various' }
  'snes'           = @{ Platform = 'Super Nintendo Entertainment System'; Note = 'Nintendo' }
  'snes-msu1'      = @{ Platform = 'Super NES CD-ROM/SNES MSU-1'; Note = 'Nintendo' }
  'socrates'       = @{ Platform = 'Socrates'; Note = 'VTech' }
  'solarus'        = @{ Platform = 'Solarus'; Note = 'Port' }
  'sonic-mania'    = @{ Platform = 'Sonic Mania'; Note = 'Port' }
  'sonic3-air'     = @{ Platform = 'Sonic 3 Angel Island Revisited'; Note = 'Port' }
  'sonicretro'     = @{ Platform = 'Star Engine/Sonic Retro Engine'; Note = 'Port' }
  'spectravideo'   = @{ Platform = 'Spectravideo'; Note = 'Spectravideo' }
  'steam'          = @{ Platform = 'Steam'; Note = 'Valve' }
  'sufami'         = @{ Platform = 'SuFami Turbo'; Note = 'Bandai' }
  'superbroswar'   = @{ Platform = 'Super Mario War'; Note = 'Port' }
  'supergrafx'     = @{ Platform = 'PC Engine/SuperGrafx/PC Engine 2'; Note = 'NEC' }
  'supervision'    = @{ Platform = 'Watara Supervision'; Note = 'Watara' }
  'supracan'       = @{ Platform = 'Super A''Can'; Note = 'Funtech Entertainment' }
  'switch'         = @{ Platform = 'Switch'; Note = 'Nintendo' }
  'systemsp'       = @{ Platform = 'Sega System SP'; Note = 'Sega' }
  'theforceengine' = @{ Platform = 'The Force Engine (Star Wars: Dark Forces)'; Note = 'Port' }
  'thextech'       = @{ Platform = 'TheXTech (Mega man)'; Note = 'Sinclair' }
  'thomson'        = @{ Platform = 'Thomson MO/TO Series Computer'; Note = 'Thomson' }
  'ti99'           = @{ Platform = 'TI-99/4/4A'; Note = 'Texas Instruments' }
  'tic80'          = @{ Platform = 'TIC-80 Fantasy Console'; Note = 'Vadim Grigoruk' }
  'traider1'       = @{ Platform = 'TR1X (Tomb Raider 1)'; Note = 'Port' }
  'traider2'       = @{ Platform = 'TR2X (Tomb Rauder 2)'; Note = 'Port' }
  'triforce'       = @{ Platform = 'Triforce'; Note = 'Namco/Sega/Nintendo' }
  'tutor'          = @{ Platform = 'Tomy Tutor/Pyuta/Grandstand Tutor'; Note = 'Tomy' }
  'tyrain'         = @{ Platform = 'TyrQuake (Quake)'; Note = 'Port' }
  'tyrquake'       = @{ Platform = 'TyrQuake (Quake 1)'; Note = 'Port' }
  'uqm'            = @{ Platform = 'The Ur-Quan Master (Star Control II)'; Note = 'Port' }
  'uzebox'         = @{ Platform = 'Uzebox Open-Source Console'; Note = 'Alec Bourque' }
  'vectrex'        = @{ Platform = 'Vectrex'; Note = 'Milton Bradley' }
  'vc4000'         = @{ Platform = 'Video Computer 4000'; Note = 'Interton' }
  'vgmplay'        = @{ Platform = 'MAME Video Game Music Player'; Note = 'Various' }
  'vircon32'       = @{ Platform = 'Vircon32 virtual console'; Note = 'Carra' }
  'vis'            = @{ Platform = 'Video Information System'; Note = 'Tandy/Memorex' }
  'vitaquake2'     = @{ Platform = 'PlayStation Vita port of Quake II'; Note = 'Port' }
  'virtualboy'     = @{ Platform = 'Virtual Boy'; Note = 'Nintendo' }
  'vpinball'       = @{ Platform = 'Visual Pinball'; Note = 'Port' }
  'voxatron'       = @{ Platform = 'Voxatron Fantasy Console'; Note = 'Lexaloffle Games' }
  'vsmile'         = @{ Platform = 'V.Smile (TV LEARNING SYSTEM)'; Note = 'VTech' }
  'wasm4'          = @{ Platform = 'WASM4 Fantasy Console'; Note = 'Aduros' }
  'wii'            = @{ Platform = 'Wii'; Note = 'Nintendo' }
  'wiiu'           = @{ Platform = 'Wii U'; Note = 'Nintendo' }
  'windows'        = @{ Platform = 'WINE'; Note = 'Bob Amstadt/Alexandre Julliard' }
  'wswan'          = @{ Platform = 'WonderSwan'; Note = 'Bandai' }
  'wswanc'         = @{ Platform = 'WonderSwan Color'; Note = 'Bandai' }
  'x1'             = @{ Platform = 'Sharp X1'; Note = 'Sharp' }
  'x68000'         = @{ Platform = 'Sharp X68000'; Note = 'Sharp' }
  'xash3d_fwgs'    = @{ Platform = 'Xash3D FWGS (Valve Games)'; Note = 'Port' }
  'xbox'           = @{ Platform = 'Xbox'; Note = 'Microsoft' }
  'xbox360'        = @{ Platform = 'Xbox 360'; Note = 'Microsoft' }
  'xegs'           = @{ Platform = 'Atari XEGS'; Note = 'Atari' }
  'xrick'          = @{ Platform = 'Rick Dangerous'; Note = 'Port' }
  'zx81'           = @{ Platform = 'Sinclair ZX81'; Note = 'Sinclair' }
  'zxspectrum'     = @{ Platform = 'ZX Spectrum'; Note = 'Sinclair' }
}

# ==================================================================================================
# FUNCTIONS
# ==================================================================================================

# --- FUNCTION: Get-PlatformInfo ---
# PURPOSE:
# - Translate a platform folder name (e.g., "psx") into a friendly platform name AND a Note.
# NOTES:
# - Falls back to returning the folder name as Platform and "" as Note if no translation is present.
function Get-PlatformInfo {
  param([string]$PlatformFolder)

  if ($PlatformMap.ContainsKey($PlatformFolder)) {
    $m = $PlatformMap[$PlatformFolder]
    return [pscustomobject]@{
      PlatformName = [string]$m.Platform
      Note         = [string]$m.Note
    }
  }

  return [pscustomobject]@{
    PlatformName = [string]$PlatformFolder
    Note         = ''
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
      ForEach-Object { if ($null -eq $_) { '' } else { ([string]$_).Trim() } } |
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
        # - Prefer resolvedName (human-friendly and stable)
        # - Fall back to path when name is missing to avoid unrelated collisions
        $groupKey     = if (-not [string]::IsNullOrWhiteSpace($resolvedName)) { $resolvedName } else { [string]$path }

        $out += [pscustomobject]@{
          PlatformFolder = $PlatformFolder
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
    $groupKey     = if (-not [string]::IsNullOrWhiteSpace($resolvedName)) { $resolvedName } else { [string]$path }

    $out += [pscustomobject]@{
      PlatformFolder = $PlatformFolder
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

  $platformFolder = [string]$t.PlatformFolder
  $platformInfo   = Get-PlatformInfo $platformFolder
  $platformName   = [string]$platformInfo.PlatformName
  $platformNote   = [string]$platformInfo.Note

  # Read and normalize all <game> entries for the platform (normal parse or salvage mode).
  $entries = @(Read-Gamelist $platformFolder $t.GamelistPath)
  if ($entries.Count -eq 0) {
    if ($isRomsRootMode) {
      Write-Host "No entries found (skipping)." -ForegroundColor Yellow
      Write-Host "Finished platform: $($t.PlatformFolder)" -ForegroundColor Cyan
    }
    continue
  }

  # Group entries into "sets" so multi-disk collections can be inferred from:
  # - .m3u visible entries (Multi-M3U)
  # - multiple entries sharing the group key where additional disks are hidden (Multi-XML)
  foreach ($group in @($entries | Group-Object GroupKey)) {

    $items        = @($group.Group)
    $hiddenItems  = @($items | Where-Object { $_.Hidden })
    $visibleItems = @($items | Where-Object { -not $_.Hidden })

    # Defensive behavior:
    # - If every entry in a group is hidden, still emit one row (otherwise the group disappears from the report).
    if ($visibleItems.Count -eq 0 -and $items.Count -gt 0) {
      $visibleItems = @($items | Select-Object -First 1)
    }

    foreach ($g in $visibleItems) {

      $pathStr = [string]$g.PathRaw
      $nameStr = [string]$g.NameResolved

      # Determine multi-disk behavior based on:
      # - .m3u path (Multi-M3U)
      # - presence of hidden sibling entries within the group (Multi-XML)
      $entryType = 'Single'
      if ($pathStr -match '(?i)\.m3u$') {
        $entryType = 'Multi-M3U'
      }
      elseif ($hiddenItems.Count -gt 0) {
        $entryType = 'Multi-XML'
      }

      $title = Get-TitleForOutput -ResolvedName $nameStr

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

        $diskCount = Get-M3UDiskCount -M3UPath $m3uFullPath
      }
      elseif ($entryType -eq 'Multi-XML') {
        $diskCount = [int]$items.Count
      }
      else {
        $diskCount = 1
      }

      # Build output row in the exact column order expected for the CSV export
      $rows += [pscustomobject]@{
        Title          = $title
        PlatformName   = $platformName
        Note           = $platformNote
        EntryType      = [string]$entryType
        DiskCount      = [int]$diskCount
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
