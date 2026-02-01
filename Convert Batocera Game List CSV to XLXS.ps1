<#
PURPOSE: Converts the CSV exported by the "Export Batocera Game List.ps1" script into an Excel .xlsx workbook
VERSION: 1.1
AUTHOR: Devin Kelley, Distant Thunderworks LLC

USER PROMPTS:
  1) Split into separate worksheets per PlatformFolder? (Yes/No)
  2) Choose CSV file to convert (GUI dialog when possible; console fallback otherwise)
  3) Choose where to save XLSX (GUI dialog when possible; console fallback otherwise)

OUTPUT:
  - Saves an .xlsx workbook to the user-selected location.
  - Defaults to the CSV folder and the same base filename with .xlsx.

FORMATTING PER SHEET:
  - Header row: Bold
  - Title column: Text format + left aligned
  - DiskCount column: centered
  - AutoFilter enabled
  - AutoFit columns
  - Freeze top row enabled (single worksheet only)
  - Selects A1 so the workbook opens without a large selection

REQUIRES:
  - Microsoft Excel installed (desktop)
  - If Excel isn't detected the script will abort
#>

# ---------------------------------------------
# Parameters (script configuration)
# ---------------------------------------------
[CmdletBinding()]
param(
    # Column name used to split into multiple worksheets (when split mode is selected)
    [string]$SplitByColumn = "PlatformFolder",

    # Worksheet name used in single-sheet mode
    [string]$SingleSheetName = "Data",

    # Formatting toggle: enable AutoFilter on each sheet
    [switch]$AutoFilter   = $true,

    # Formatting toggle: AutoFit used columns on each sheet
    [switch]$AutoSize     = $true,

    # Formatting toggle: freeze top row (single-sheet mode only)
    [switch]$FreezeTopRow = $true
)

# -------------------------------------------------------------------------------------------------
# Verify Microsoft Excel is installed before doing anything else
# -------------------------------------------------------------------------------------------------
try {
    $testExcel = New-Object -ComObject Excel.Application
    $testExcel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($testExcel)
} catch {
    Write-Host ""
    Write-Host "Microsoft Excel does not appear to be installed on this system." -ForegroundColor Red
    Write-Host "This script requires the desktop version of Excel to generate XLSX files." -ForegroundColor Red
    Write-Host ""
    return
}

# =================================================================================================
# SECTION: UI mode helpers (GUI when STA + WinForms available, console fallback otherwise)
# =================================================================================================

$script:WinFormsOk = $false

function Test-IsSTA {
    # Detect whether the current PowerShell thread is running in STA mode (needed for WinForms dialogs)
    try { return [System.Threading.Thread]::CurrentThread.ApartmentState -eq 'STA' }
    catch { return $false }
}

function Ensure-WinForms {
    # Load WinForms assemblies used by message boxes and file dialogs
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing       | Out-Null

    # Enable WinForms visual styles (non-fatal if unavailable)
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}
}

function Can-UseGui {
    return ((Test-IsSTA) -and $script:WinFormsOk)
}

function Show-TopMostMessageBox {
    <#
      Shows a message box with a hidden top-most owner form so it stays in front of editors/ISE.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Text,
        [Parameter(Mandatory=$true)][string]$Title,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    # Create a tiny off-screen top-most owner window for proper z-order
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $owner.StartPosition = 'Manual'
    $owner.Size = New-Object System.Drawing.Size(1,1)
    $owner.Location = New-Object System.Drawing.Point(-32000,-32000)
    $owner.ShowInTaskbar = $false
    $owner.Show() | Out-Null

    try {
        # Show the message box owned by the hidden form
        return [System.Windows.Forms.MessageBox]::Show($owner, $Text, $Title, $Buttons, $Icon)
    } finally {
        # Clean up the owner form
        $owner.Close()
        $owner.Dispose()
    }
}

function Ask-SplitMode {
    <#
      Prompts the user to choose split mode:
        - Yes: one worksheet per unique value in $SplitByColumn
        - No:  single worksheet
    #>
    param([string]$SplitColumnName)

    $msg = "Split into separate worksheets by '$SplitColumnName'?" + [Environment]::NewLine +
           "Yes = one tab per value" + [Environment]::NewLine +
           "No  = all rows in a single tab"

    $title = "CSV → XLSX Options"

    Write-Host ""
    Write-Host "Prompt: choose split mode..." -ForegroundColor Cyan

    if (Can-UseGui) {
        $result = Show-TopMostMessageBox -Text $msg -Title $title `
            -Buttons ([System.Windows.Forms.MessageBoxButtons]::YesNo) `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Question)

        return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
    } else {
        Write-Host $msg
        $ans = Read-Host "Type Y for Yes (split) or N for No (single sheet)"
        return ($ans -match '^(?i)y')
    }
}

function Pick-CsvFile {
    <#
      Prompts the user to select a CSV file:
        - GUI OpenFileDialog when STA + WinForms available
        - Console prompt otherwise
    #>
    Write-Host ""
    Write-Host "Prompt: choose CSV file..." -ForegroundColor Cyan

    if (Can-UseGui) {
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.StartPosition = 'Manual'
        $owner.Size = New-Object System.Drawing.Size(1,1)
        $owner.Location = New-Object System.Drawing.Point(-32000,-32000)
        $owner.ShowInTaskbar = $false
        $owner.Show() | Out-Null

        try {
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = "Select CSV file to convert"
            $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            $dlg.Multiselect = $false
            $dlg.CheckFileExists = $true
            $dlg.RestoreDirectory = $true

            if ($dlg.ShowDialog($owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
            return $dlg.FileName
        } finally {
            $owner.Close()
            $owner.Dispose()
        }
    } else {
        $p = Read-Host "Enter full path to the CSV file"
        if ([string]::IsNullOrWhiteSpace($p)) { return $null }
        return $p
    }
}

function Pick-XlsxSavePath {
    <#
      Prompts the user to select the output .xlsx save path:
        - Starts in InitialDirectory when possible
        - Pre-fills SuggestedFileName
        - Adds .xlsx extension when omitted
        - GUI SaveFileDialog when STA + WinForms available
        - Console prompt otherwise
    #>
    param(
        [Parameter(Mandatory=$true)][string]$InitialDirectory,
        [Parameter(Mandatory=$true)][string]$SuggestedFileName
    )

    Write-Host ""
    Write-Host "Prompt: choose XLSX save location..." -ForegroundColor Cyan

    if (Can-UseGui) {
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.StartPosition = 'Manual'
        $owner.Size = New-Object System.Drawing.Size(1,1)
        $owner.Location = New-Object System.Drawing.Point(-32000,-32000)
        $owner.ShowInTaskbar = $false
        $owner.Show() | Out-Null

        try {
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Title = "Save XLSX as..."
            $dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            $dlg.OverwritePrompt = $true
            $dlg.RestoreDirectory = $true
            $dlg.AddExtension = $true
            $dlg.DefaultExt   = "xlsx"

            if (-not [string]::IsNullOrWhiteSpace($InitialDirectory) -and (Test-Path -LiteralPath $InitialDirectory)) {
                $dlg.InitialDirectory = $InitialDirectory
            }

            $dlg.FileName = $SuggestedFileName

            if ($dlg.ShowDialog($owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
            return $dlg.FileName
        } finally {
            $owner.Close()
            $owner.Dispose()
        }
    } else {
        Write-Host "Suggested output: $(Join-Path $InitialDirectory $SuggestedFileName)"
        $p = Read-Host "Enter full path where to save the XLSX (or press Enter to accept suggested)"

        if ([string]::IsNullOrWhiteSpace($p)) {
            $p = (Join-Path $InitialDirectory $SuggestedFileName)
        }

        if ([System.IO.Path]::GetExtension($p) -eq "") { $p = $p + ".xlsx" }
        return $p
    }
}

# =================================================================================================
# SECTION: Freeze panes helper (single worksheet top row freeze)
# =================================================================================================

function Freeze-WorksheetTopRow {
    <#
      Freezes the top row by selecting A2 and enabling FreezePanes on the workbook window.
      This relies on an available Window object for the workbook.
    #>
    param([Parameter(Mandatory=$true)]$Worksheet)

    try {
        $wb = $Worksheet.Parent
        $null = $wb.Activate()
        $null = $Worksheet.Activate()

        $wnd = $null
        try { $wnd = $wb.Windows.Item(1) } catch {}

        if ($null -eq $wnd) { return }

        $wnd.FreezePanes  = $false
        $wnd.SplitRow     = 0
        $wnd.SplitColumn  = 0

        $null = $Worksheet.Range("A2").Select()
        $wnd.FreezePanes = $true
        $null = $Worksheet.Range("A1").Select()
    } catch {
        # Ignore freeze failures
    }
}

# =================================================================================================
# SECTION: Sheet name utilities (Excel-safe name + uniqueness)
# =================================================================================================

function Sanitize-SheetName {
    <#
      Converts a string into an Excel-safe worksheet name:
        - Replaces invalid characters
        - Trims whitespace
        - Enforces maximum length (31)
    #>
    param([Parameter(Mandatory=$true)][string]$Name)

    $n = $Name
    $n = $n -replace '[:\\\/\?\*\[\]]', '_'
    $n = $n.Trim()

    if ([string]::IsNullOrWhiteSpace($n)) { $n = "Sheet" }

    $n = $n.TrimEnd("'")
    if ($n.Length -gt 31) { $n = $n.Substring(0,31) }

    return $n
}

function Get-UniqueSheetName {
    <#
      Ensures a worksheet name is unique within the workbook by suffixing " (n)" as needed.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][hashtable]$Used
    )

    $base = Sanitize-SheetName $Name
    $candidate = $base

    if (-not $Used.ContainsKey($candidate)) {
        $Used[$candidate] = $true
        return $candidate
    }

    $i = 2
    while ($true) {
        $suffix = " ($i)"
        $maxBaseLen = 31 - $suffix.Length

        $trimBase = $base
        if ($trimBase.Length -gt $maxBaseLen) { $trimBase = $trimBase.Substring(0, $maxBaseLen) }

        $candidate = $trimBase + $suffix

        if (-not $Used.ContainsKey($candidate)) {
            $Used[$candidate] = $true
            return $candidate
        }

        $i++
    }
}

# =================================================================================================
# SECTION: Excel COM constants + COM cleanup helper
# =================================================================================================

$xlLeft   = -4131
$xlCenter = -4108

$xlCalculationManual    = -4135
$xlCalculationAutomatic = -4105

$xlOpenXMLWorkbook      = 51

# Paste constants (used already; keep explicit)
$xlPasteValues  = -4163
$xlPasteFormats = -4122

# SpecialCells
$xlCellTypeVisible = 12

# OpenText constants
$xlDelimited = 1
$xlTextQualifierDoubleQuote = 1

function Release-ComObject {
    <#
      Releases COM references to reduce Excel.exe process leaks.
    #>
    param([Parameter(Mandatory=$true)]$Obj)

    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Obj) } catch {}
}

# =================================================================================================
# SECTION: Worksheet formatting utilities
# =================================================================================================

function Get-HeaderMap {
    <#
      Reads worksheet row 1 and returns a hashtable mapping header text -> column index.
    #>
    param([Parameter(Mandatory=$true)]$Worksheet)

    $map = @{}

    $used = $null
    try { $used = $Worksheet.UsedRange } catch {}
    if ($null -eq $used) { return $map }

    $colCount = 0
    try { $colCount = $used.Columns.Count } catch { $colCount = 0 }

    for ($c = 1; $c -le $colCount; $c++) {
        $h = ""
        try { $h = [string]$Worksheet.Cells.Item(1, $c).Text } catch { $h = "" }

        if (-not [string]::IsNullOrWhiteSpace($h)) {
            $map[$h] = $c
        }
    }

    Release-ComObject $used
    return $map
}

function Apply-WorksheetFormatting {
    <#
      Applies formatting to a single worksheet:
        - Bold header row
        - Formats specific columns by header name
        - Enables AutoFilter
        - AutoFits columns
        - Freezes top row (when allowed)
        - Selects A1
    #>
    param(
        [Parameter(Mandatory=$true)]$Worksheet,
        [switch]$AllowFreezeTopRow = $true
    )

    $used = $null
    try { $used = $Worksheet.UsedRange } catch {}
    if ($null -eq $used) { return }

    $rowCount = 0
    $colCount = 0
    try { $rowCount = $used.Rows.Count } catch { $rowCount = 0 }
    try { $colCount = $used.Columns.Count } catch { $colCount = 0 }
    if ($rowCount -lt 1 -or $colCount -lt 1) { Release-ComObject $used; return }

    $hdr = Get-HeaderMap -Worksheet $Worksheet

    try { $used.Rows.Item(1).Font.Bold = $true } catch {}

    # Title column formatting
    if ($hdr.ContainsKey("Title") -and $rowCount -ge 2) {
        $rng = $null
        try {
            $c = [int]$hdr["Title"]
            $rng = $Worksheet.Range($Worksheet.Cells.Item(2,$c), $Worksheet.Cells.Item($rowCount,$c))
            $rng.NumberFormat = "@"
            $rng.HorizontalAlignment = $xlLeft
        } catch {} finally {
            if ($rng) { Release-ComObject $rng }
        }
    }

    # DiskCount column formatting
    if ($hdr.ContainsKey("DiskCount") -and $rowCount -ge 2) {
        $rng = $null
        try {
            $c = [int]$hdr["DiskCount"]
            $rng = $Worksheet.Range($Worksheet.Cells.Item(2,$c), $Worksheet.Cells.Item($rowCount,$c))
            $rng.HorizontalAlignment = $xlCenter
        } catch {} finally {
            if ($rng) { Release-ComObject $rng }
        }
    }

    if ($AutoFilter) { try { $null = $used.AutoFilter() } catch {} }
    if ($AutoSize)   { try { $null = $used.Columns.AutoFit() } catch {} }

    if ($AllowFreezeTopRow -and $FreezeTopRow) {
        Freeze-WorksheetTopRow -Worksheet $Worksheet
    }

    try {
        $null = $Worksheet.Activate()
        $null = $Worksheet.Range("A1").Select()
    } catch {}

    Release-ComObject $used
}

# =================================================================================================
# SECTION: CSV open helper (more locale-stable; falls back to Open)
# =================================================================================================

function Open-CsvWorkbook {
    <#
      Tries to open CSV using OpenText with an explicit comma delimiter (reduces locale delimiter issues).
      Falls back to Workbooks.Open on failure.
    #>
    param(
        [Parameter(Mandatory=$true)]$ExcelApp,
        [Parameter(Mandatory=$true)][string]$CsvPath
    )

    # Try OpenText first
    try {
        $ExcelApp.Workbooks.OpenText(
            $CsvPath,            # Filename
            65001,               # Origin (attempt UTF-8 code page)
            $null,               # StartRow
            $xlDelimited,        # DataType
            $xlTextQualifierDoubleQuote, # TextQualifier
            $false,              # ConsecutiveDelimiter
            $false,              # Tab
            $false,              # Semicolon
            $true,               # Comma
            $false,              # Space
            $false               # Other
        ) | Out-Null

        # When OpenText is used, Excel typically sets the opened workbook as ActiveWorkbook
        return $ExcelApp.ActiveWorkbook
    } catch {
        # Fallback: classic Open (locale dependent)
        return $ExcelApp.Workbooks.Open($CsvPath)
    }
}

# =================================================================================================
# SECTION: Split key normalization helpers
# =================================================================================================

function Normalize-ExcelTextForKey {
    <#
      Normalizes Excel cell .Text for stable key grouping & filtering:
        - Converts NBSP to normal space
        - Trims standard whitespace
      IMPORTANT: We use the same normalized value both for key discovery and filtering, to reduce mismatches.
    #>
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) { return "" }

    # Replace non-breaking space with regular space
    $t = $Text -replace [char]0x00A0, ' '

    # Trim regular whitespace
    return $t.Trim()
}

# =================================================================================================
# SECTION: Main workflow
# =================================================================================================

# Load WinForms when possible (do NOT swallow and then attempt GUI types)
try {
    Ensure-WinForms
    $script:WinFormsOk = $true
} catch {
    $script:WinFormsOk = $false
}

$splitMode = Ask-SplitMode -SplitColumnName $SplitByColumn
$csvPath = Pick-CsvFile

if ([string]::IsNullOrWhiteSpace($csvPath)) {
    Write-Host "Cancelled (no CSV selected)." -ForegroundColor Yellow
    return
}

if (-not (Test-Path -LiteralPath $csvPath)) {
    throw "CSV file not found: $csvPath"
}

$outDir = Split-Path -Parent $csvPath
$outBase = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
$suggestedXlsxName = $outBase + ".xlsx"

$xlsxPath = Pick-XlsxSavePath -InitialDirectory $outDir -SuggestedFileName $suggestedXlsxName

if ([string]::IsNullOrWhiteSpace($xlsxPath)) {
    Write-Host "Cancelled (no XLSX save path chosen)." -ForegroundColor Yellow
    return
}

# Validate output directory exists (console mode can produce non-existent dirs)
$outSaveDir = Split-Path -Parent $xlsxPath
if ([string]::IsNullOrWhiteSpace($outSaveDir) -or -not (Test-Path -LiteralPath $outSaveDir)) {
    throw "Output directory not found: $outSaveDir"
}

Write-Host ""
Write-Host "Output will be written to:" -ForegroundColor Cyan
Write-Host "  $xlsxPath" -ForegroundColor Cyan
Write-Host ""

$excel  = $null
$wbCsv  = $null
$wbOut  = $null
$oldCalc = $null

# Extra COM refs we want to explicitly release
$wsSrc  = $null
$usedSrc = $null
$visible = $null

try {
    Write-Host "Launching Excel (COM)..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    try {
        $oldCalc = $excel.Calculation
        $excel.Calculation = $xlCalculationManual
    } catch {}

    Write-Host "Opening CSV in Excel..." -ForegroundColor Cyan
    $wbCsv = Open-CsvWorkbook -ExcelApp $excel -CsvPath $csvPath

    $wsSrc = $wbCsv.Worksheets.Item(1)
    $usedSrc = $wsSrc.UsedRange

    if ($null -eq $usedSrc -or $usedSrc.Rows.Count -lt 1) {
        $wbOut = $excel.Workbooks.Add()
        $wsOut = $wbOut.Worksheets.Item(1)
        $wsOut.Name = Sanitize-SheetName $SingleSheetName
        Apply-WorksheetFormatting -Worksheet $wsOut -AllowFreezeTopRow:$true

        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }
        $wbOut.SaveAs($xlsxPath, $xlOpenXMLWorkbook)

        Write-Host "Wrote XLSX (empty input): $xlsxPath" -ForegroundColor Green
        return
    }

    $hdrMap = Get-HeaderMap -Worksheet $wsSrc

    if ($splitMode -and -not $hdrMap.ContainsKey($SplitByColumn)) {
        throw "Split column '$SplitByColumn' not found in CSV headers."
    }

    if (-not $splitMode) {
        # ------------------------------
        # Single-sheet mode
        # ------------------------------
        Write-Host "Single-sheet mode: formatting and saving..." -ForegroundColor Cyan

        $wsSrc.Name = Sanitize-SheetName $SingleSheetName
        Apply-WorksheetFormatting -Worksheet $wsSrc -AllowFreezeTopRow:$true

        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }
        $wbCsv.SaveAs($xlsxPath, $xlOpenXMLWorkbook)
    }
    else {
        # ------------------------------
        # Split mode
        # ------------------------------
        Write-Host "Split mode: creating destination workbook..." -ForegroundColor Cyan

        $wbOut = $excel.Workbooks.Add()

        $usedNames = @{}
        $createdSheetNames = New-Object System.Collections.Generic.List[string]

        $splitColIndex = [int]$hdrMap[$SplitByColumn]

        $rowCount = 0
        try { $rowCount = $usedSrc.Rows.Count } catch { $rowCount = 0 }

        $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        if ($rowCount -ge 2) {
            for ($r = 2; $r -le $rowCount; $r++) {
                $cellText = ""
                try { $cellText = [string]$wsSrc.Cells.Item($r, $splitColIndex).Text } catch { $cellText = "" }

                $v = Normalize-ExcelTextForKey $cellText
                if ([string]::IsNullOrWhiteSpace($v)) { $v = "(blank)" }

                [void]$values.Add($v)
            }
        } else {
            [void]$values.Add("(blank)")
        }

        $uniqueKeys = @($values) | Sort-Object

        # Enable AutoFilter on the source used range (when supported)
        try { $null = $usedSrc.AutoFilter() } catch {}

        foreach ($key in $uniqueKeys) {
            $sheetName = Get-UniqueSheetName -Name $key -Used $usedNames
            Write-Host "Writing worksheet: $sheetName" -ForegroundColor Cyan

            $field = $splitColIndex - $usedSrc.Column + 1

            # Reset filter state (best effort)
            try {
                $wsSrc.AutoFilterMode = $false
                $null = $usedSrc.AutoFilter()
            } catch {}

            # Apply filter for current key (best effort; blanks are tricky across locales)
            $filterOk = $true
            try {
                if ($key -eq "(blank)") {
                    # Prefer empty-string criteria for blanks
                    $null = $usedSrc.AutoFilter($field, "")
                } else {
                    $null = $usedSrc.AutoFilter($field, $key)
                }
            } catch {
                $filterOk = $false
            }

            if (-not $filterOk) {
                Write-Host "  Skipped (filter failed for key '$key')." -ForegroundColor Yellow
                Write-Host ""
                continue
            }

            # Copy visible (filtered) cells — SpecialCells throws if none visible
            $visible = $null
            try {
                $visible = $usedSrc.SpecialCells($xlCellTypeVisible)
            } catch {
                Write-Host "  Skipped (no visible rows for key '$key')." -ForegroundColor Yellow
                Write-Host ""
                continue
            }

            try {
                $null = $visible.Copy()

                $wsDest = $wbOut.Worksheets.Add()
                $wsDest.Name = $sheetName

                $null = $wsDest.Range("A1").PasteSpecial($xlPasteValues)
                $null = $wsDest.Range("A1").PasteSpecial($xlPasteFormats)

                # Formatting, but do not attempt freeze panes in split mode
                Apply-WorksheetFormatting -Worksheet $wsDest -AllowFreezeTopRow:$false

                [void]$createdSheetNames.Add($sheetName)

                Write-Host "Completed" -ForegroundColor White
                Write-Host ""

                Release-ComObject $wsDest
                $wsDest = $null
            } finally {
                if ($visible) { Release-ComObject $visible; $visible = $null }
            }
        }

        # Reorder worksheets alphabetically using names (avoid holding worksheet COM objects)
        try {
            $sortedNames = $createdSheetNames.ToArray() | Sort-Object

            for ($i = $sortedNames.Count - 1; $i -ge 0; $i--) {
                $nm = $sortedNames[$i]
                try {
                    $ws = $wbOut.Worksheets.Item($nm)
                    $ws.Move($wbOut.Worksheets.Item(1)) | Out-Null
                    Release-ComObject $ws
                } catch {}
            }
        } catch {}

        # Remove default workbook sheets not created by us (keep at least one sheet)
        for ($i = $wbOut.Worksheets.Count; $i -ge 1; $i--) {
            $ws = $wbOut.Worksheets.Item($i)
            $nm = [string]$ws.Name

            if (-not $usedNames.ContainsKey($nm) -and $wbOut.Worksheets.Count -gt 1) {
                try { $null = $ws.Delete() } catch { try { $ws.Delete() | Out-Null } catch {} }
            }

            Release-ComObject $ws
        }

        Write-Host "Saving XLSX..." -ForegroundColor Cyan

        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }
        $wbOut.SaveAs($xlsxPath, $xlOpenXMLWorkbook)
    }

    Write-Host ""
    Write-Host "Wrote XLSX: $xlsxPath" -ForegroundColor Green
}
finally {
    # Restore Excel calculation mode when available
    if ($excel -and $null -ne $oldCalc) {
        try { $excel.Calculation = $oldCalc } catch {}
    } elseif ($excel) {
        try { $excel.Calculation = $xlCalculationAutomatic } catch {}
    }

    # Release extra COM references if held
    if ($visible) { try { Release-ComObject $visible } catch {}; $visible = $null }
    if ($usedSrc) { try { Release-ComObject $usedSrc } catch {}; $usedSrc = $null }
    if ($wsSrc)   { try { Release-ComObject $wsSrc }   catch {}; $wsSrc   = $null }

    # Close destination workbook without prompts
    if ($wbOut) {
        try { $wbOut.Close($false) } catch {}
        Release-ComObject $wbOut
        $wbOut = $null
    }

    # Close source workbook without prompts
    if ($wbCsv) {
        try { $wbCsv.Close($false) } catch {}
        Release-ComObject $wbCsv
        $wbCsv = $null
    }

    # Quit Excel application
    if ($excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
        $excel = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
