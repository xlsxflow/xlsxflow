# Generates reader test fixtures with real Microsoft Excel (COM automation, Windows only).
# Usage: powershell -ExecutionPolicy Bypass -File scripts/fixtures/make-excel.ps1
# Output: test/fixtures/excel-*.xlsx (then run make-expected.py to refresh oracles)
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot '..\..\test\fixtures'
New-Item -ItemType Directory -Force $out | Out-Null
$out = (Resolve-Path $out).Path

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

# Formatting/layout commands can be refused by an unlicensed (reduced-functionality) Excel.
# They are best-effort: expectations are derived from the saved file, not from this script.
function Try-Step([scriptblock]$step, [string]$what) {
  try { & $step } catch { Write-Warning "skipped ${what}: $($_.Exception.Message)" }
}

function New-Book {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 5) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  return $wb
}

function Fill-Book($wb) {
  # --- Types: every value kind, sparse cells, awkward strings ---
  $s = $wb.Worksheets.Item(1); $s.Name = 'Types'
  $s.Range('A1').Value2 = 'text'
  $s.Range('B1').Value2 = 42
  $s.Range('C1').Value2 = -3.5
  $s.Range('D1').Value2 = 1e20
  $s.Range('E1').Value2 = 1e-10
  $s.Range('F1').Formula = '=0.1+0.2'
  $s.Range('A2').Value2 = $true
  $s.Range('B2').Value2 = $false
  $s.Range('C2').Formula = '=1/0'
  $s.Range('D2').Formula = '=NA()'
  $s.Range('E2').Formula = '=B1*2'
  $s.Range('F2').Formula = '="a"&"b"'
  $s.Range('G2').Formula = '=B1>1'
  $s.Range('A4').Value2 = 'row 4 after a gap'
  $s.Range('Z4').Value2 = 'far column'
  # Built from code points: PowerShell 5.1 reads BOM-less scripts as ANSI
  $u = { param($cps) -join ($cps | ForEach-Object { [char]::ConvertFromUtf32($_) }) }
  $s.Range('A5').Value2 = 'emoji ' + (& $u @(0x1F600)) + ' arabic ' + (& $u @(0x645,0x631,0x62D,0x628,0x627)) + ' cjk ' + (& $u @(0x6F22,0x5B57))
  $s.Range('B5').Value2 = '  padded  '
  $s.Range('C5').Value2 = 'x & y < z > "q" ''s'''
  $s.Range('D5').Value2 = "line1`nline2"
  $s.Range('E5').NumberFormat = '@'
  $s.Range('E5').Value2 = '00123'
  $s.Range('F5').Value2 = ('L' * 32767)
  $s.Range('G5').Value2 = 'text'   # duplicate shared string
  $s.Range('A6').Value2 = '_x0041_ literal escape-looking text'

  # --- Dates: date formats vs. non-date number formats ---
  $d = $wb.Worksheets.Item(2); $d.Name = 'Dates'
  $d.Range('A1').Formula = '=DATE(2024,2,29)';          $d.Range('A1').NumberFormat = 'yyyy-mm-dd'
  $d.Range('B1').Value2 = 45351.75;                      $d.Range('B1').NumberFormat = 'dd/mm/yyyy hh:mm'
  $d.Range('C1').Value2 = 0.5;                           $d.Range('C1').NumberFormat = 'hh:mm:ss'
  $d.Range('D1').Value2 = 45000;                         $d.Range('D1').NumberFormat = '[$-409]mmmm d, yyyy;@'
  $d.Range('E1').Value2 = 1;                             $d.Range('E1').NumberFormat = 'm/d/yyyy'
  $d.Range('F1').Value2 = 61;                            $d.Range('F1').NumberFormat = 'm/d/yyyy'
  $d.Range('A2').Value2 = 0.25;                          $d.Range('A2').NumberFormat = '0.00%'
  $d.Range('B2').Value2 = 1234.5;                        $d.Range('B2').NumberFormat = '#,##0.00 "USD"'
  $d.Range('C2').Value2 = 1.5;                           $d.Range('C2').NumberFormat = '[h]:mm'
  $d.Range('D2').Value2 = 7;                             $d.Range('D2').NumberFormat = '0 "days"'
  $d.Range('E2').Value2 = 123.456;                       $d.Range('E2').NumberFormat = '[Red]0.0;[Blue]-0.0'

  # --- Rich text & layout features ---
  $r = $wb.Worksheets.Item(3); $r.Name = 'Rich & Layout'
  $r.Range('A1').Value2 = 'bold then plain'
  Try-Step { $r.Range('A1').Characters(1, 4).Font.Bold = $true } 'rich text'
  $r.Range('A2').Value2 = 'merged'
  Try-Step { [void]$r.Range('A2:C3').Merge() } 'merge'
  $r.Range('A5').Value2 = 'hidden row'
  Try-Step { $r.Rows(5).Hidden = $true } 'hidden row'
  $r.Range('D1').Value2 = 'hidden column'
  Try-Step { $r.Columns(4).Hidden = $true } 'hidden column'
  Try-Step { $r.Activate(); $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.FreezePanes = $true } 'freeze panes'

  # --- Bulk: 20k x 5 formula cells ---
  $b = $wb.Worksheets.Item(4); $b.Name = 'Bulk'
  $rows = 20000
  # Formulas with cached values: exercises the reader on 100k <f>+<v> cells, shared strings via t="str"
  $b.Range("A1:A$rows").Formula = '=ROW()-1'
  $b.Range("B1:B$rows").Formula = '="name "&MOD(ROW()-1,500)'
  $b.Range("C1:C$rows").Formula = '=(ROW()-1)*0.5'
  $b.Range("D1:D$rows").Formula = '=MOD(ROW()-1,2)=0'
  $b.Range("E1:E$rows").Formula = '="unique-"&(ROW()-1)'

  # --- Empty sheet (Excel writes <sheetData/>) ---
  $e = $wb.Worksheets.Item(5); Try-Step { $e.Name = 'Empty' } 'rename empty sheet'

  # Tab order differs from part order: the last-created sheet becomes the first tab
  Try-Step { $b.Move($wb.Worksheets.Item(1)) } 'sheet reorder'
  $wb.Worksheets.Item(1).Activate()
}

try {
  $wb = New-Book; Fill-Book $wb
  $wb.SaveAs((Join-Path $out 'excel-basic.xlsx'), 51)          # xlOpenXMLWorkbook
  $wb.SaveAs((Join-Path $out 'excel-strict.xlsx'), 61)         # xlOpenXMLStrictWorkbook
  $wb.Close($false)

  $wb = New-Book; $wb.Date1904 = $true; Fill-Book $wb
  $wb.SaveAs((Join-Path $out 'excel-1904.xlsx'), 51)
  $wb.Close($false)
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()   # lets EXCEL.EXE actually exit
}
Write-Output "Excel fixtures written to $out"
