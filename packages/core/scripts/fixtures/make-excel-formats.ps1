# Builds one workbook in desktop Excel and saves it as .xlsx, .xls and .ods (test/fixtures/excel-made.*).
# Usage: powershell -ExecutionPolicy Bypass -File scripts/fixtures/make-excel-formats.ps1
param([string]$Out = (Resolve-Path (Join-Path $PSScriptRoot '..\..\test\fixtures')).Path)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'Data'

  $ws.Range('A1').Value2 = 'Text'
  $ws.Range('B1').Value2 = 42
  $ws.Range('C1').Value2 = 3.14159
  $ws.Range('D1').Value2 = 46303            # 2026-10-08
  $ws.Range('D1').NumberFormat = 'yyyy-mm-dd'
  $ws.Range('E1').Value2 = $true
  $ws.Range('F1').Formula = '=1/0'
  $ws.Range('G1').Formula = '=B1*2'
  $ws.Range('H1').Formula = '=A1&"!"'
  $ws.Range('I1').Value2 = "h$([char]0xE9)llo $([char]0x2014) $([char]0x65E5)$([char]0x672C)"
  $ws.Range('J1').Value2 = 1234567890123
  $ws.Range('K1').Value2 = -0.5
  $ws.Range('L1').Value2 = 0.256
  $ws.Range('L1').NumberFormat = '0.0%'
  $ws.Range('M1').Value2 = 46303.586458333333   # 14:04:30
  $ws.Range('M1').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
  $ws.Range('N1').Formula = '=B1>40'
  $ws.Range('O1').Formula = '=""'
  $ws.Range('P1').Value2 = 12.5
  $ws.Range('P1').NumberFormat = '#,##0.00 "USD"'
  $ws.Range('Q1').Value2 = 100
  $ws.Range('R1').Formula = '=NA()'

  $ws.Range('A2').Value2 = 'hidden row'
  $ws.Rows(2).Hidden = $true
  $ws.Range('S1').Value2 = 'hidden column'
  $ws.Columns('S').Hidden = $true

  $ws.Range('A3').Value2 = 'Merged'
  $ws.Range('A3:C3').Merge()

  # One string far longer than a BIFF record, mixing 8-bit and 16-bit characters, so it splits
  $sb = New-Object System.Text.StringBuilder
  for ($i = 0; $i -lt 1500; $i++) { [void]$sb.Append("abc$i "); if ($i % 100 -eq 0) { [void]$sb.Append([char]0x3042) } }
  $ws.Range('A5').Value2 = $sb.ToString()

  # Enough distinct strings that the shared string table spans many CONTINUE records
  $n = 3000
  $target = $ws.Range("A10:C$($n + 9)")
  $ws.Range("A10:A$($n + 9)").Formula = '="String number "&(ROW()-10)'
  $ws.Range("B10:B$($n + 9)").Formula = '=IF(MOD(ROW()-10,7)=0,"Unicode "&UNICHAR(1046)&(ROW()-10),"plain "&(ROW()-10))'
  $ws.Range("C10:C$($n + 9)").Formula = '=(ROW()-10)*1.5'
  [void]$target.Copy()
  [void]$target.PasteSpecial(-4163)   # values only
  $ws.Range('E10').Formula = "=SUM(C10:C$($n + 9))"

  $ws.Activate()
  $xl.ActiveWindow.SplitRow = 1
  $xl.ActiveWindow.SplitColumn = 1
  $xl.ActiveWindow.FreezePanes = $true

  $s2 = $wb.Worksheets.Item(2); $s2.Name = 'Second'
  $s2.Range('A1').Value2 = 'second sheet'
  $s2.Range('B2').Value2 = 7
  $s3 = $wb.Worksheets.Item(3); $s3.Name = 'Secret'
  $s3.Range('A1').Value2 = 'hidden sheet'
  $s3.Visible = 0   # xlSheetHidden

  $ws.Activate()
  $wb.SaveAs("$Out\excel-made.xlsx", 51)
  $wb.SaveAs("$Out\excel-made.xls", 56)
  $wb.SaveAs("$Out\excel-made.ods", 60)
  $wb.Close($false)

  'ok'
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
