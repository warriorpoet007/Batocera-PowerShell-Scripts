<#
PURPOSE: Converts the CSV exported by the "Export Batocera Game List.ps1" script into an Excel .xlsx workbook
VERSION: 1.0
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
  - Freeze top row enabled (single worksheet only; split-mode behavior depends on Excel window context)
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

    # Formatting toggle: freeze top row (only works when writing a single-worksheet file)
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
# SECTION: UI mode helpers (GUI when STA, console fallback otherwise)
# =================================================================================================

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

    # Build the prompt text
    $msg = "Split into separate worksheets by '$SplitColumnName'?" + [Environment]::NewLine +
           "Yes = one tab per value" + [Environment]::NewLine +
           "No  = all rows in a single tab"

    # Set the dialog title
    $title = "CSV → XLSX Options"

    # Write a visible cue in the console
    Write-Host ""
    Write-Host "Prompt: choose split mode..." -ForegroundColor Cyan

    if (Test-IsSTA) {
        # GUI prompt using a top-most message box
        $result = Show-TopMostMessageBox -Text $msg -Title $title `
            -Buttons ([System.Windows.Forms.MessageBoxButtons]::YesNo) `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Question)

        # Convert DialogResult to a boolean split flag
        return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
    } else {
        # Console prompt fallback
        Write-Host $msg
        $ans = Read-Host "Type Y for Yes (split) or N for No (single sheet)"

        # Treat Y/y as split mode
        return ($ans -match '^(?i)y')
    }
}

function Pick-CsvFile {
    <#
      Prompts the user to select a CSV file:
        - GUI OpenFileDialog when STA
        - Console prompt otherwise
    #>
    Write-Host ""
    Write-Host "Prompt: choose CSV file..." -ForegroundColor Cyan

    if (Test-IsSTA) {
        # Create a tiny off-screen top-most owner window for the file dialog
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.StartPosition = 'Manual'
        $owner.Size = New-Object System.Drawing.Size(1,1)
        $owner.Location = New-Object System.Drawing.Point(-32000,-32000)
        $owner.ShowInTaskbar = $false
        $owner.Show() | Out-Null

        try {
            # Configure the OpenFileDialog
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = "Select CSV file to convert"
            $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            $dlg.Multiselect = $false
            $dlg.CheckFileExists = $true
            $dlg.RestoreDirectory = $true

            # Return $null on cancel
            if ($dlg.ShowDialog($owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

            # Return selected file path
            return $dlg.FileName
        } finally {
            # Clean up the owner form
            $owner.Close()
            $owner.Dispose()
        }
    } else {
        # Console prompt fallback
        $p = Read-Host "Enter full path to the CSV file"

        # Treat blank as cancel
        if ([string]::IsNullOrWhiteSpace($p)) { return $null }

        # Return typed path
        return $p
    }
}

function Pick-XlsxSavePath {
    <#
      Prompts the user to select the output .xlsx save path:
        - Starts in InitialDirectory when possible
        - Pre-fills SuggestedFileName
        - Adds .xlsx extension when omitted
        - GUI SaveFileDialog when STA
        - Console prompt otherwise
    #>
    param(
        [Parameter(Mandatory=$true)][string]$InitialDirectory,
        [Parameter(Mandatory=$true)][string]$SuggestedFileName
    )

    Write-Host ""
    Write-Host "Prompt: choose XLSX save location..." -ForegroundColor Cyan

    if (Test-IsSTA) {
        # Create a tiny off-screen top-most owner window for the file dialog
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true
        $owner.StartPosition = 'Manual'
        $owner.Size = New-Object System.Drawing.Size(1,1)
        $owner.Location = New-Object System.Drawing.Point(-32000,-32000)
        $owner.ShowInTaskbar = $false
        $owner.Show() | Out-Null

        try {
            # Configure the SaveFileDialog
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Title = "Save XLSX as..."
            $dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            $dlg.OverwritePrompt = $true
            $dlg.RestoreDirectory = $true
            $dlg.AddExtension = $true
            $dlg.DefaultExt   = "xlsx"

            # Set initial directory when it exists
            if (-not [string]::IsNullOrWhiteSpace($InitialDirectory) -and (Test-Path -LiteralPath $InitialDirectory)) {
                $dlg.InitialDirectory = $InitialDirectory
            }

            # Pre-fill the output filename
            $dlg.FileName = $SuggestedFileName

            # Return $null on cancel
            if ($dlg.ShowDialog($owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

            # Return selected path
            return $dlg.FileName
        } finally {
            # Clean up the owner form
            $owner.Close()
            $owner.Dispose()
        }
    } else {
        # Show suggested output path
        Write-Host "Suggested output: $(Join-Path $InitialDirectory $SuggestedFileName)"

        # Console prompt fallback
        $p = Read-Host "Enter full path where to save the XLSX (or press Enter to accept suggested)"

        # Accept suggested output on empty input
        if ([string]::IsNullOrWhiteSpace($p)) {
            return (Join-Path $InitialDirectory $SuggestedFileName)
        }

        # Append extension if missing
        if ([System.IO.Path]::GetExtension($p) -eq "") { $p = $p + ".xlsx" }

        # Return typed output path
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
        # Get the workbook containing the worksheet
        $wb = $Worksheet.Parent

        # Activate workbook and worksheet to establish context for window operations
        $null = $wb.Activate()
        $null = $Worksheet.Activate()

        # Retrieve the first window for the workbook
        $wnd = $null
        try { $wnd = $wb.Windows.Item(1) } catch {}

        # Exit if a window is not available
        if ($null -eq $wnd) { return }

        # Clear any prior freeze or split settings
        $wnd.FreezePanes  = $false
        $wnd.SplitRow     = 0
        $wnd.SplitColumn  = 0

        # Select the row below the header so Excel freezes row 1
        $null = $Worksheet.Range("A2").Select()
        $wnd.FreezePanes = $true

        # Return selection to A1
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

    # Start with the raw name
    $n = $Name

    # Replace invalid characters with underscores
    $n = $n -replace '[:\\\/\?\*\[\]]', '_'

    # Trim whitespace
    $n = $n.Trim()

    # Use a fallback if empty
    if ([string]::IsNullOrWhiteSpace($n)) { $n = "Sheet" }

    # Avoid trailing apostrophes
    $n = $n.TrimEnd("'")

    # Truncate to Excel sheet name limit
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

    # Sanitize the base name first
    $base = Sanitize-SheetName $Name
    $candidate = $base

    # Use base name if unused
    if (-not $Used.ContainsKey($candidate)) {
        $Used[$candidate] = $true
        return $candidate
    }

    # Increment suffix until a unique name is found
    $i = 2
    while ($true) {
        # Build suffix
        $suffix = " ($i)"

        # Compute the maximum base length allowed when adding the suffix
        $maxBaseLen = 31 - $suffix.Length

        # Trim base if needed
        $trimBase = $base
        if ($trimBase.Length -gt $maxBaseLen) { $trimBase = $trimBase.Substring(0, $maxBaseLen) }

        # Combine trimmed base + suffix
        $candidate = $trimBase + $suffix

        # Accept first unused candidate
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

# Excel constant: horizontal alignment (left)
$xlLeft   = -4131

# Excel constant: horizontal alignment (center)
$xlCenter = -4108

# Excel constant: calculation mode (manual)
$xlCalculationManual = -4135

# Excel constant: calculation mode (automatic)
$xlCalculationAutomatic = -4105

# Excel constant: SaveAs format ID for .xlsx
$xlOpenXMLWorkbook = 51

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

    # Initialize mapping table
    $map = @{}

    # Determine used range on the sheet
    $used = $Worksheet.UsedRange
    if ($null -eq $used) { return $map }

    # Iterate columns in the used range
    $colCount = $used.Columns.Count
    for ($c = 1; $c -le $colCount; $c++) {
        # Read cell text for the header label
        $h = [string]$Worksheet.Cells.Item(1, $c).Text

        # Store non-empty header labels
        if (-not [string]::IsNullOrWhiteSpace($h)) {
            $map[$h] = $c
        }
    }

    return $map
}

function Apply-WorksheetFormatting {
    <#
      Applies formatting to a single worksheet:
        - Bold header row
        - Formats specific columns by header name
        - Enables AutoFilter
        - AutoFits columns
        - Freezes top row (if enabled)
        - Selects A1
    #>
    param([Parameter(Mandatory=$true)]$Worksheet)

    # Determine used range
    $used = $Worksheet.UsedRange
    if ($null -eq $used) { return }

    # Read used range bounds
    $rowCount = $used.Rows.Count
    $colCount = $used.Columns.Count
    if ($rowCount -lt 1 -or $colCount -lt 1) { return }

    # Build a header->column index map for the sheet
    $hdr = Get-HeaderMap -Worksheet $Worksheet

    # Bold the header row
    try { $used.Rows.Item(1).Font.Bold = $true } catch {}

    # Apply Title column formatting (rows 2..end)
    if ($hdr.ContainsKey("Title") -and $rowCount -ge 2) {
        $c = [int]$hdr["Title"]
        $rng = $Worksheet.Range($Worksheet.Cells.Item(2,$c), $Worksheet.Cells.Item($rowCount,$c))
        $rng.NumberFormat = "@"
        $rng.HorizontalAlignment = $xlLeft
    }

    # Apply DiskCount column formatting (rows 2..end)
    if ($hdr.ContainsKey("DiskCount") -and $rowCount -ge 2) {
        $c = [int]$hdr["DiskCount"]
        $rng = $Worksheet.Range($Worksheet.Cells.Item(2,$c), $Worksheet.Cells.Item($rowCount,$c))
        $rng.HorizontalAlignment = $xlCenter
    }

    # Enable AutoFilter on the used range
    if ($AutoFilter) { try { $null = $used.AutoFilter() } catch {} }

    # AutoFit columns in the used range
    if ($AutoSize) { try { $null = $used.Columns.AutoFit() } catch {} }

    # Freeze the top row (worksheet-level call)
    if ($FreezeTopRow) { Freeze-WorksheetTopRow -Worksheet $Worksheet }

    # Select A1 to avoid an expanded selection on open
    try {
        $null = $Worksheet.Activate()
        $null = $Worksheet.Range("A1").Select()
    } catch {}
}

# =================================================================================================
# SECTION: Main workflow
# =================================================================================================

# Load WinForms when possible (non-fatal if unavailable)
try { Ensure-WinForms } catch {}

# Prompt for split mode selection
$splitMode = Ask-SplitMode -SplitColumnName $SplitByColumn

# Prompt for source CSV file selection
$csvPath = Pick-CsvFile

# Exit on cancel
if ([string]::IsNullOrWhiteSpace($csvPath)) {
    Write-Host "Cancelled (no CSV selected)." -ForegroundColor Yellow
    return
}

# Verify the CSV path exists
if (-not (Test-Path -LiteralPath $csvPath)) {
    throw "CSV file not found: $csvPath"
}

# Compute default output directory from CSV location
$outDir = Split-Path -Parent $csvPath

# Compute default base name from CSV filename
$outBase = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)

# Build suggested output XLSX name
$suggestedXlsxName = $outBase + ".xlsx"

# Prompt for final XLSX output path
$xlsxPath = Pick-XlsxSavePath -InitialDirectory $outDir -SuggestedFileName $suggestedXlsxName

# Exit on cancel
if ([string]::IsNullOrWhiteSpace($xlsxPath)) {
    Write-Host "Cancelled (no XLSX save path chosen)." -ForegroundColor Yellow
    return
}

# Echo selected output path
Write-Host ""
Write-Host "Output will be written to:" -ForegroundColor Cyan
Write-Host "  $xlsxPath" -ForegroundColor Cyan
Write-Host ""

# Initialize COM object holders for cleanup
$excel = $null
$wbCsv = $null
$wbOut = $null
$oldCalc = $null

try {
    # Create Excel Application COM instance
    Write-Host "Launching Excel (COM)..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application

    # Set Excel automation flags
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    # Switch calculation to manual for faster bulk operations
    try {
        $oldCalc = $excel.Calculation
        $excel.Calculation = $xlCalculationManual
    } catch {}

    # Open the CSV as a workbook
    Write-Host "Opening CSV in Excel..." -ForegroundColor Cyan
    $wbCsv = $excel.Workbooks.Open($csvPath)

    # Read the first worksheet from the CSV workbook
    $wsSrc = $wbCsv.Worksheets.Item(1)

    # Read the used range for the CSV worksheet
    $usedSrc = $wsSrc.UsedRange

    # Handle empty CSV by saving an empty workbook
    if ($null -eq $usedSrc -or $usedSrc.Rows.Count -lt 1) {
        $wbOut = $excel.Workbooks.Add()
        $wsOut = $wbOut.Worksheets.Item(1)
        $wsOut.Name = Sanitize-SheetName $SingleSheetName
        Apply-WorksheetFormatting -Worksheet $wsOut

        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }
        $wbOut.SaveAs($xlsxPath, $xlOpenXMLWorkbook)

        Write-Host "Wrote XLSX (empty input): $xlsxPath" -ForegroundColor Green
        return
    }

    # Build header map for split column checks
    $hdrMap = Get-HeaderMap -Worksheet $wsSrc

    # Validate that split column exists when split mode is enabled
    if ($splitMode -and -not $hdrMap.ContainsKey($SplitByColumn)) {
        throw "Split column '$SplitByColumn' not found in CSV headers."
    }

    if (-not $splitMode) {
        # ------------------------------
        # Single-sheet mode
        # ------------------------------

        Write-Host "Single-sheet mode: formatting and saving..." -ForegroundColor Cyan

        # Rename the sheet
        $wsSrc.Name = Sanitize-SheetName $SingleSheetName

        # Apply per-sheet formatting
        Apply-WorksheetFormatting -Worksheet $wsSrc

        # Replace output file if it exists
        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }

        # Save the CSV workbook as XLSX
        $wbCsv.SaveAs($xlsxPath, $xlOpenXMLWorkbook)

    } else {
        # ------------------------------
        # Split mode: one sheet per PlatformFolder value
        # ------------------------------

        Write-Host "Split mode: creating destination workbook..." -ForegroundColor Cyan

        # Create destination workbook
        $wbOut = $excel.Workbooks.Add()

        # Track used sheet names to avoid duplicates
        $usedNames = @{}

        # Find the column index of the split column
        $splitColIndex = [int]$hdrMap[$SplitByColumn]

        # Read the row count for unique-value discovery
        $rowCount = $usedSrc.Rows.Count

        # Collect unique split keys
        $values = [System.Collections.Generic.HashSet[string]]::new()

        # Gather unique values (using Value2 for stable comparisons)
        if ($rowCount -ge 2) {
            for ($r = 2; $r -le $rowCount; $r++) {
                $raw = $wsSrc.Cells.Item($r, $splitColIndex).Value2
                $v = if ($null -eq $raw) { "" } else { [string]$raw }
                $v = $v.Trim()
                if ([string]::IsNullOrWhiteSpace($v)) { $v = "(blank)" }
                [void]$values.Add($v)
            }
        } else {
            [void]$values.Add("(blank)")
        }

        # Materialize unique keys as a sorted list
        $uniqueKeys = @($values) | Sort-Object

        # Enable AutoFilter on the source used range (when supported)
        try { $null = $usedSrc.AutoFilter() } catch {}

        foreach ($key in $uniqueKeys) {
            # Build a unique, Excel-safe worksheet name
            $sheetName = Get-UniqueSheetName -Name $key -Used $usedNames

            # Announce the current output sheet
            Write-Host "Writing worksheet: $sheetName" -ForegroundColor Cyan

            # Compute the filter field index relative to the used range start column
            $field = $splitColIndex - $usedSrc.Column + 1

            # Clear existing filters
            try {
                $wsSrc.AutoFilterMode = $false
                $null = $usedSrc.AutoFilter()
            } catch {}

            # Apply filter for current key
            if ($key -eq "(blank)") {
                $null = $usedSrc.AutoFilter($field, "=")
            } else {
                $null = $usedSrc.AutoFilter($field, $key)
            }

            # Copy visible (filtered) cells
            $visible = $usedSrc.SpecialCells(12)  # 12 = xlCellTypeVisible
            $null = $visible.Copy()

            # Add destination worksheet
            $wsDest = $wbOut.Worksheets.Add()
            $wsDest.Name = $sheetName

            # Paste values into destination
            $null = $wsDest.Range("A1").PasteSpecial(-4163)  # -4163 = xlPasteValues

            # Paste formats into destination
            $null = $wsDest.Range("A1").PasteSpecial(-4122)  # -4122 = xlPasteFormats

            # Apply per-sheet formatting to destination
            Apply-WorksheetFormatting -Worksheet $wsDest

            # Emit per-sheet completion status
            Write-Host "Completed" -ForegroundColor White
            Write-Host ""

            # Release per-iteration COM references
            Release-ComObject $wsDest
        }

        # Reorder worksheets alphabetically (left to right)
        try {
            $sorted = @()

            # Build a list of only the sheets we created (names in $usedNames)
            foreach ($ws in @($wbOut.Worksheets)) {
                $nm = [string]$ws.Name
                if ($usedNames.ContainsKey($nm)) { $sorted += $ws }
            }

            # Sort by name (case-insensitive)
            $sorted = $sorted | Sort-Object -Property Name

            # Move in reverse so the final order matches the sorted list
            for ($i = $sorted.Count - 1; $i -ge 0; $i--) {
                $sorted[$i].Move($wbOut.Worksheets.Item(1)) | Out-Null
            }
        } catch {}

        # Remove default workbook sheets that were not renamed to a used name
        for ($i = $wbOut.Worksheets.Count; $i -ge 1; $i--) {
            $ws = $wbOut.Worksheets.Item($i)
            $nm = [string]$ws.Name

            # Delete sheets not present in the used-name set (while keeping at least one sheet)
            if (-not $usedNames.ContainsKey($nm) -and $wbOut.Worksheets.Count -gt 1) {
                try { $null = $ws.Delete() } catch { $ws.Delete() | Out-Null }
            }

            # Release worksheet COM reference
            Release-ComObject $ws
        }

        # Save the destination workbook
        Write-Host "Saving XLSX..." -ForegroundColor Cyan

        # Replace output file if it exists
        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }

        # Save as .xlsx
        $wbOut.SaveAs($xlsxPath, $xlOpenXMLWorkbook)
    }

    # Print final success message
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

    # Close destination workbook without prompts
    if ($wbOut) {
        try { $wbOut.Close($false) } catch {}
        Release-ComObject $wbOut
    }

    # Close source workbook without prompts
    if ($wbCsv) {
        try { $wbCsv.Close($false) } catch {}
        Release-ComObject $wbCsv
    }

    # Quit Excel application
    if ($excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }

    # Force garbage collection for COM cleanup
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

